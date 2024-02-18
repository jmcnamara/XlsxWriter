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


class TestWriteDefinedNames(unittest.TestCase):
    """
    Test the Workbook _write_defined_names() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_defined_names_1(self):
        """Test the _write_defined_names() method"""

        self.workbook.defined_names = [["_xlnm.Print_Titles", 0, "Sheet1!$1:$1", 0]]

        self.workbook._write_defined_names()

        exp = """<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName></definedNames>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_defined_names_2(self):
        """Test the _write_defined_names() method"""

        self.workbook.add_worksheet()
        self.workbook.add_worksheet()
        self.workbook.add_worksheet("Sheet 3")

        self.workbook.define_name("""'Sheet 3'!Bar""", """='Sheet 3'!$A$1""")
        self.workbook.define_name("""Abc""", """=Sheet1!$A$1""")
        self.workbook.define_name("""Baz""", """=0.98""")
        self.workbook.define_name("""Sheet1!Bar""", """=Sheet1!$A$1""")
        self.workbook.define_name("""Sheet2!Bar""", """=Sheet2!$A$1""")
        self.workbook.define_name("""Sheet2!aaa""", """=Sheet2!$A$1""")
        self.workbook.define_name("""'Sheet 3'!car""", '="Saab 900"')
        self.workbook.define_name("""_Egg""", """=Sheet1!$A$1""")
        self.workbook.define_name("""_Fog""", """=Sheet1!$A$1""")

        self.workbook._prepare_defined_names()
        self.workbook._write_defined_names()

        exp = """<definedNames><definedName name="_Egg">Sheet1!$A$1</definedName><definedName name="_Fog">Sheet1!$A$1</definedName><definedName name="aaa" localSheetId="1">Sheet2!$A$1</definedName><definedName name="Abc">Sheet1!$A$1</definedName><definedName name="Bar" localSheetId="2">'Sheet 3'!$A$1</definedName><definedName name="Bar" localSheetId="0">Sheet1!$A$1</definedName><definedName name="Bar" localSheetId="1">Sheet2!$A$1</definedName><definedName name="Baz">0.98</definedName><definedName name="car" localSheetId="2">"Saab 900"</definedName></definedNames>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
