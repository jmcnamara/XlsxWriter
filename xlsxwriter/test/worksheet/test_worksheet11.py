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

from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """

    def test_assemble_xml_file(self):
        """Test writing formula-like strings with strings_to_formulas off."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.str_table = SharedStringTable()

        worksheet.select()

        # Neither string should be converted to a formula.
        worksheet.strings_to_formulas = False
        worksheet.write("A1", "=1+2")
        worksheet.write("A2", "{=1+2}")

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="A1:A2"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <sheetData>
                    <row r="1" spans="1:1">
                      <c r="A1" t="s">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="2" spans="1:1">
                      <c r="A2" t="s">
                        <v>1</v>
                      </c>
                    </row>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
