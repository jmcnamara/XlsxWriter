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

        worksheet.write("A1", 1)
        worksheet.write("A2", 2)
        worksheet.write("A3", 3)
        worksheet.write("A4", 4)
        worksheet.write("A5", 5)
        worksheet.write("A6", 6)
        worksheet.write("A7", 7)
        worksheet.write("A8", 8)
        worksheet.write("A9", 9)

        worksheet.write("A12", 75)

        worksheet.conditional_format(
            "A1", {"type": "icon_set", "icon_style": "3_arrows", "reverse_icons": True}
        )

        worksheet.conditional_format(
            "A2", {"type": "icon_set", "icon_style": "3_flags", "icons_only": True}
        )

        worksheet.conditional_format(
            "A3",
            {
                "type": "icon_set",
                "icon_style": "3_traffic_lights_rimmed",
                "icons_only": True,
                "reverse_icons": True,
            },
        )

        worksheet.conditional_format(
            "A4",
            {
                "type": "icon_set",
                "icon_style": "3_symbols_circled",
                "icons": [
                    {"value": 80},
                    {"value": 20},
                ],
            },
        )

        worksheet.conditional_format(
            "A5",
            {
                "type": "icon_set",
                "icon_style": "4_arrows",
                "icons": [
                    {"criteria": ">"},
                    {"criteria": ">"},
                    {"criteria": ">"},
                ],
            },
        )

        worksheet.conditional_format(
            "A6",
            {
                "type": "icon_set",
                "icon_style": "4_red_to_black",
                "icons": [
                    {"criteria": ">=", "type": "number", "value": 90},
                    {"criteria": "<", "type": "percentile", "value": 50},
                    {"criteria": "<=", "type": "percent", "value": 25},
                ],
            },
        )

        worksheet.conditional_format(
            "A7",
            {
                "type": "icon_set",
                "icon_style": "4_traffic_lights",
                "icons": [{"value": "=$A$12"}],
            },
        )

        worksheet.conditional_format(
            "A8",
            {
                "type": "icon_set",
                "icon_style": "5_arrows_gray",
                "icons": [{"type": "formula", "value": "=$A$12"}],
            },
        )

        worksheet.conditional_format(
            "A9",
            {
                "type": "icon_set",
                "icon_style": "5_quarters",
                "icons": [
                    {"type": "percentile", "value": 70},
                    {"type": "percentile", "value": 50},
                    {"type": "percentile", "value": 30},
                    {"type": "percentile", "value": 10},
                    {"type": "percentile", "value": -1},
                ],
                "reverse_icons": True,
            },
        )

        worksheet._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="A1:A12"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <sheetData>
                    <row r="1" spans="1:1">
                      <c r="A1">
                        <v>1</v>
                      </c>
                    </row>
                    <row r="2" spans="1:1">
                      <c r="A2">
                        <v>2</v>
                      </c>
                    </row>
                    <row r="3" spans="1:1">
                      <c r="A3">
                        <v>3</v>
                      </c>
                    </row>
                    <row r="4" spans="1:1">
                      <c r="A4">
                        <v>4</v>
                      </c>
                    </row>
                    <row r="5" spans="1:1">
                      <c r="A5">
                        <v>5</v>
                      </c>
                    </row>
                    <row r="6" spans="1:1">
                      <c r="A6">
                        <v>6</v>
                      </c>
                    </row>
                    <row r="7" spans="1:1">
                      <c r="A7">
                        <v>7</v>
                      </c>
                    </row>
                    <row r="8" spans="1:1">
                      <c r="A8">
                        <v>8</v>
                      </c>
                    </row>
                    <row r="9" spans="1:1">
                      <c r="A9">
                        <v>9</v>
                      </c>
                    </row>
                    <row r="12" spans="1:1">
                      <c r="A12">
                        <v>75</v>
                      </c>
                    </row>
                  </sheetData>
                  <conditionalFormatting sqref="A1">
                    <cfRule type="iconSet" priority="1">
                      <iconSet iconSet="3Arrows" reverse="1">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="33"/>
                        <cfvo type="percent" val="67"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A2">
                    <cfRule type="iconSet" priority="2">
                      <iconSet iconSet="3Flags" showValue="0">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="33"/>
                        <cfvo type="percent" val="67"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A3">
                    <cfRule type="iconSet" priority="3">
                      <iconSet iconSet="3TrafficLights2" showValue="0" reverse="1">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="33"/>
                        <cfvo type="percent" val="67"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A4">
                    <cfRule type="iconSet" priority="4">
                      <iconSet iconSet="3Symbols">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="20"/>
                        <cfvo type="percent" val="80"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A5">
                    <cfRule type="iconSet" priority="5">
                      <iconSet iconSet="4Arrows">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="25" gte="0"/>
                        <cfvo type="percent" val="50" gte="0"/>
                        <cfvo type="percent" val="75" gte="0"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A6">
                    <cfRule type="iconSet" priority="6">
                      <iconSet iconSet="4RedToBlack">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="25"/>
                        <cfvo type="percentile" val="50"/>
                        <cfvo type="num" val="90"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A7">
                    <cfRule type="iconSet" priority="7">
                      <iconSet iconSet="4TrafficLights">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="25"/>
                        <cfvo type="percent" val="50"/>
                        <cfvo type="percent" val="$A$12"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A8">
                    <cfRule type="iconSet" priority="8">
                      <iconSet iconSet="5ArrowsGray">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="20"/>
                        <cfvo type="percent" val="40"/>
                        <cfvo type="percent" val="60"/>
                        <cfvo type="formula" val="$A$12"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A9">
                    <cfRule type="iconSet" priority="9">
                      <iconSet iconSet="5Quarters" reverse="1">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percentile" val="10"/>
                        <cfvo type="percentile" val="30"/>
                        <cfvo type="percentile" val="50"/>
                        <cfvo type="percentile" val="70"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
