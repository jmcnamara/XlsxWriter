###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...workbook import Workbook


class TestWriteCalcPr(unittest.TestCase):
    """
    Test the Workbook _write_calc_pr() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_calc_pr(self):
        """Test the _write_calc_pr() method."""

        self.workbook._write_calc_pr()

        exp = """<calcPr calcId="124519" fullCalcOnLoad="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_calc_mode_auto_except_tables(self):
        """
        Test the _write_calc_pr() method with the calculation mode set
        to auto_except_tables.

        """

        self.workbook.set_calc_mode("auto_except_tables")
        self.workbook._write_calc_pr()

        exp = """<calcPr calcId="124519" calcMode="autoNoTable" fullCalcOnLoad="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_calc_mode_manual(self):
        """
        Test the _write_calc_pr() method with the calculation mode set to
        manual.

        """

        self.workbook.set_calc_mode("manual")
        self.workbook._write_calc_pr()

        exp = """<calcPr calcId="124519" calcMode="manual" calcOnSave="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_calc_pr2(self):
        """Test the _write_calc_pr() method with non-default calc id."""

        self.workbook.set_calc_mode("auto", 12345)
        self.workbook._write_calc_pr()

        exp = """<calcPr calcId="12345" fullCalcOnLoad="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
