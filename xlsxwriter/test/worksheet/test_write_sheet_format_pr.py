###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...worksheet import Worksheet


class TestWriteSheetFormatPr(unittest.TestCase):
    """
    Test the Worksheet _write_sheet_format_pr() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_sheet_format_pr(self):
        """Test the _write_sheet_format_pr() method"""

        self.worksheet._write_sheet_format_pr()

        exp = """<sheetFormatPr defaultRowHeight="15"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
