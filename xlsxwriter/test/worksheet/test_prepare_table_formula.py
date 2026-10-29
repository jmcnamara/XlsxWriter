###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.worksheet import Worksheet


class TestPrepareTableFormula(unittest.TestCase):
    """
    Test the _prepare_table_formula Worksheet method for different formula types.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)
        self.prepare = self.worksheet._prepare_table_formula

    def test_prepare_table_formula_empty(self):
        self.assertEqual(self.prepare(""), "")

    def test_prepare_table_formula_this_row_shortcut(self):
        self.assertEqual(self.prepare("@"), "[#This Row],")

    def test_prepare_table_formula_at_inside_string_is_unchanged(self):
        self.assertEqual(self.prepare('"Comment @"'), '"Comment @"')

    def test_prepare_table_formula_column_range(self):
        got = self.prepare("SUM(Table1[@[Column1]:[Column3]])")
        expected = "SUM(Table1[[#This Row],[Column1]:[Column3]])"
        self.assertEqual(got, expected)

    def test_prepare_table_formula_hyperlink_with_at_in_url(self):
        formula = (
            "HYPERLINK(CONCAT(\"http://myweb.com:1677/'@md=d&path/to/sour/...'@/\","
            '[@CL],"?ac=10"),[@CL])'
        )
        expected = (
            "HYPERLINK(CONCAT(\"http://myweb.com:1677/'@md=d&path/to/sour/...'@/\","
            '[[#This Row],CL],"?ac=10"),[[#This Row],CL])'
        )
        self.assertEqual(self.prepare(formula), expected)
