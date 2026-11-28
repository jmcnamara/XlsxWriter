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

from xlsxwriter.chart import Chart


class TestWriteADefRPr(unittest.TestCase):
    """
    Test the Chart _write_a_def_rpr() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.chart = Chart()
        self.chart._set_filehandle(self.fh)

    def test_write_a_def_rpr_fractional_size(self):
        """Test the _write_a_def_rpr() method with a fractional font size"""

        font = self.chart._convert_font_args({"size": 1.13})
        self.chart._write_a_def_rpr(font)

        exp = """<a:defRPr sz="113" baseline="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
