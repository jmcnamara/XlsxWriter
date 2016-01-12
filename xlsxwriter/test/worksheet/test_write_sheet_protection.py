###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteSheetProtection(unittest.TestCase):
    """
    Test the Worksheet _write_sheet_protection() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_sheet_protection_1(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_2(self):
        """Test the _write_sheet_protection() method."""

        password = 'password'
        options = {}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection password="83AF" sheet="1" objects="1" scenarios="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_3(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'select_locked_cells': 0}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_4(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'format_cells': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_5(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'format_columns': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_6(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'format_rows': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" formatRows="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_7(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'insert_columns': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" insertColumns="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_8(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'insert_rows': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" insertRows="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_9(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'insert_hyperlinks': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" insertHyperlinks="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_10(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'delete_columns': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" deleteColumns="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_11(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'delete_rows': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" deleteRows="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_12(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'sort': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" sort="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_13(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'autofilter': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_14(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'pivot_tables': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" pivotTables="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_15(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'objects': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" scenarios="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_16(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'scenarios': 1}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_17(self):
        """Test the _write_sheet_protection() method."""

        password = ''
        options = {'format_cells': 1, 'select_locked_cells': 0, 'select_unlocked_cells': 0}

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1" selectUnlockedCells="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_18(self):
        """Test the _write_sheet_protection() method."""

        password = 'drowssap'
        options = {
            'objects': 1,
            'scenarios': 1,
            'format_cells': 1,
            'format_columns': 1,
            'format_rows': 1,
            'insert_columns': 1,
            'insert_rows': 1,
            'insert_hyperlinks': 1,
            'delete_columns': 1,
            'delete_rows': 1,
            'select_locked_cells': 0,
            'sort': 1,
            'autofilter': 1,
            'pivot_tables': 1,
            'select_unlocked_cells': 0,
        }

        self.worksheet.protect(password, options)
        self.worksheet._write_sheet_protection()

        exp = """<sheetProtection password="996B" sheet="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" selectLockedCells="1" sort="0" autoFilter="0" pivotTables="0" selectUnlockedCells="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
