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


class TestWriteSheetData(unittest.TestCase):
    """
    Test the Worksheet _write_sheet_data() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_sheet_data(self):
        """Test the _write_sheet_data() method"""

        self.worksheet._write_sheet_data()

        exp = """<sheetData/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
