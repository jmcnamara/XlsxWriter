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


class TestWriteDefinedName(unittest.TestCase):
    """
    Test the Workbook _write_defined_name() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_defined_name(self):
        """Test the _write_defined_name() method"""

        self.workbook._write_defined_name(['_xlnm.Print_Titles', 0, 'Sheet1!$1:$1', 0])

        exp = """<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
