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


class TestWriteSheets(unittest.TestCase):
    """
    Test the Workbook _write_sheets() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_sheets(self):
        """Test the _write_sheets() method"""

        self.workbook.add_worksheet("Sheet2")
        self.workbook._write_sheets()

        exp = """<sheets><sheet name="Sheet2" sheetId="1" r:id="rId1"/></sheets>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
