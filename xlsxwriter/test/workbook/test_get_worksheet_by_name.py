###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ..helperfunctions import _xml_to_list
from ...workbook import Workbook


class TestAssembleWorkbook(unittest.TestCase):
    """
    Test assembling a complete Workbook file.

    """
    def test_get_worksheet_by_name01(self):
        """Test get_worksheet_by_name()"""
        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        exp = workbook.add_worksheet()
        got = workbook.get_worksheet_by_name('Sheet1')
        workbook.fileclosed = 1

        self.assertEqual(got, exp)

    def test_get_worksheet_by_name02(self):
        """Test get_worksheet_by_name()"""
        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        workbook.add_worksheet()
        exp = workbook.add_worksheet()
        got = workbook.get_worksheet_by_name('Sheet2')
        workbook.fileclosed = 1

        self.assertEqual(got, exp)

    def test_get_worksheet_by_name03(self):
        """Test get_worksheet_by_name()"""
        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        exp = workbook.add_worksheet('Sheet 3')
        got = workbook.get_worksheet_by_name('Sheet 3')
        workbook.fileclosed = 1

        self.assertEqual(got, exp)

    def test_get_worksheet_by_name04(self):
        """Test get_worksheet_by_name()"""
        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        exp = workbook.add_worksheet("Sheet '4")
        got = workbook.get_worksheet_by_name("Sheet '4")
        workbook.fileclosed = 1

        self.assertEqual(got, exp)

    def test_get_worksheet_by_name05(self):
        """Test get_worksheet_by_name()"""
        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        exp = None
        got = workbook.get_worksheet_by_name("Sheet 5")
        workbook.fileclosed = 1

        self.assertEqual(got, exp)
