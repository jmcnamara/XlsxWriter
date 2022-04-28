###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...workbook import Workbook


class TestWriteWorkbookPr(unittest.TestCase):
    """
    Test the Workbook _write_workbook_pr() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_workbook_pr(self):
        """Test the _write_workbook_pr() method"""

        self.workbook._write_workbook_pr()

        exp = """<workbookPr defaultThemeVersion="124226"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
