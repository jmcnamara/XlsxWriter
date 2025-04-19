###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """

    def test_assemble_xml_file(self):
        """Test writing a worksheet with no cell data."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()
        worksheet.name = "Sheet1"
        worksheet.excel_version = 2010

        data = [-2, 2, 3, -1, 0]
        worksheet.write_row("A1", data)
        worksheet.write_row("A2", data)
        worksheet.write_row("A3", data)
        worksheet.write_row("A4", data)
        worksheet.write_row("A5", data)
        worksheet.write_row("A6", data)
        worksheet.write_row("A7", data)

        # Set up sparklines.
        worksheet.add_sparkline(
            "F1", {"range": "A1:E1", "type": "column", "high_point": 1}
        )

        worksheet.add_sparkline(
            "F2", {"range": "A2:E2", "type": "column", "low_point": 1}
        )

        worksheet.add_sparkline(
            "F3", {"range": "A3:E3", "type": "column", "negative_points": 1}
        )

        worksheet.add_sparkline(
            "F4", {"range": "A4:E4", "type": "column", "first_point": 1}
        )

        worksheet.add_sparkline(
            "F5", {"range": "A5:E5", "type": "column", "last_point": 1}
        )

        worksheet.add_sparkline(
            "F6", {"range": "A6:E6", "type": "column", "markers": 1}
        )

        worksheet.add_sparkline(
            "F7",
            {
                "range": "A7:E7",
                "type": "column",
                "high_point": 1,
                "low_point": 1,
                "negative_points": 1,
                "first_point": 1,
                "last_point": 1,
                "markers": 1,
            },
        )

        worksheet._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                  <dimension ref="A1:E7"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
                  <sheetData>
                    <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A1">
                        <v>-2</v>
                      </c>
                      <c r="B1">
                        <v>2</v>
                      </c>
                      <c r="C1">
                        <v>3</v>
                      </c>
                      <c r="D1">
                        <v>-1</v>
                      </c>
                      <c r="E1">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="2" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A2">
                        <v>-2</v>
                      </c>
                      <c r="B2">
                        <v>2</v>
                      </c>
                      <c r="C2">
                        <v>3</v>
                      </c>
                      <c r="D2">
                        <v>-1</v>
                      </c>
                      <c r="E2">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="3" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A3">
                        <v>-2</v>
                      </c>
                      <c r="B3">
                        <v>2</v>
                      </c>
                      <c r="C3">
                        <v>3</v>
                      </c>
                      <c r="D3">
                        <v>-1</v>
                      </c>
                      <c r="E3">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="4" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A4">
                        <v>-2</v>
                      </c>
                      <c r="B4">
                        <v>2</v>
                      </c>
                      <c r="C4">
                        <v>3</v>
                      </c>
                      <c r="D4">
                        <v>-1</v>
                      </c>
                      <c r="E4">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="5" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A5">
                        <v>-2</v>
                      </c>
                      <c r="B5">
                        <v>2</v>
                      </c>
                      <c r="C5">
                        <v>3</v>
                      </c>
                      <c r="D5">
                        <v>-1</v>
                      </c>
                      <c r="E5">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="6" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A6">
                        <v>-2</v>
                      </c>
                      <c r="B6">
                        <v>2</v>
                      </c>
                      <c r="C6">
                        <v>3</v>
                      </c>
                      <c r="D6">
                        <v>-1</v>
                      </c>
                      <c r="E6">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="7" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A7">
                        <v>-2</v>
                      </c>
                      <c r="B7">
                        <v>2</v>
                      </c>
                      <c r="C7">
                        <v>3</v>
                      </c>
                      <c r="D7">
                        <v>-1</v>
                      </c>
                      <c r="E7">
                        <v>0</v>
                      </c>
                    </row>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                        <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" markers="1" high="1" low="1" first="1" last="1" negative="1">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A7:E7</xm:f>
                              <xm:sqref>F7</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" markers="1">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A6:E6</xm:f>
                              <xm:sqref>F6</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" last="1">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A5:E5</xm:f>
                              <xm:sqref>F5</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" first="1">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A4:E4</xm:f>
                              <xm:sqref>F4</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" negative="1">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A3:E3</xm:f>
                              <xm:sqref>F3</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" low="1">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A2:E2</xm:f>
                              <xm:sqref>F2</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" high="1">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A1:E1</xm:f>
                              <xm:sqref>F1</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                      </x14:sparklineGroups>
                    </ext>
                  </extLst>
                </worksheet>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
