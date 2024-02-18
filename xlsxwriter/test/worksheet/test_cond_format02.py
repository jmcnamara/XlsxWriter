###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ..helperfunctions import _xml_to_list
from ...worksheet import Worksheet


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """

    def test_assemble_xml_file(self):
        """Test writing a worksheet with conditional formatting."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()

        worksheet.write("A1", 10)
        worksheet.write("A2", 20)
        worksheet.write("A3", 30)
        worksheet.write("A4", 40)

        worksheet.write("B1", 5)

        worksheet.conditional_format(
            "A1:A1",
            {
                "type": "cell",
                "format": None,
                "criteria": "greater than",
                "value": "$B$1",
            },
        )

        worksheet._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="A1:B4"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <sheetData>
                    <row r="1" spans="1:2">
                      <c r="A1">
                        <v>10</v>
                      </c>
                      <c r="B1">
                        <v>5</v>
                      </c>
                    </row>
                    <row r="2" spans="1:2">
                      <c r="A2">
                        <v>20</v>
                      </c>
                    </row>
                    <row r="3" spans="1:2">
                      <c r="A3">
                        <v>30</v>
                      </c>
                    </row>
                    <row r="4" spans="1:2">
                      <c r="A4">
                        <v>40</v>
                      </c>
                    </row>
                  </sheetData>
                  <conditionalFormatting sqref="A1">
                    <cfRule type="cellIs" priority="1" operator="greaterThan">
                      <formula>$B$1</formula>
                    </cfRule>
                  </conditionalFormatting>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
