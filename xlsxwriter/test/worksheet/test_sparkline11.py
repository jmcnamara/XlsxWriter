###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ..helperfunctions import _xml_to_list
from ...worksheet import Worksheet


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
        worksheet.name = 'Sheet1'
        worksheet.excel_version = 2010

        data = [-2, 2, 3, -1, 0]
        worksheet.write_row('A1', data)
        worksheet.write_row('A2', data)
        worksheet.write_row('A3', data)
        worksheet.write_row('A4', [1, 2, 3, 4, 5])

        # Set up sparklines.
        worksheet.add_sparkline('F1', {'range': 'A1:E1',
                                       'max': 0.5,
                                       'min': -0.5,
                                       'axis': True,
                                       'reverse': True,
                                       'empty_cells': 'zero',
                                       'weight': 0.25,
                                       'high_point': True,
                                       'low_point': True,
                                       'negative_points': True,
                                       'first_point': True,
                                       'last_point': True,
                                       'markers': True,
                                       })

        worksheet.add_sparkline('F2', {'range': 'A2:E2',
                                       'max': 'group',
                                       'min': 'group',
                                       'empty_cells': 'connect',
                                       'weight': 2.25,
                                       })

        worksheet.add_sparkline('F3', {'range': 'A3:E3',
                                       'max': 'group',
                                       'min': '0',
                                       'show_hidden': True,
                                       'weight': 6,
                                       'date_axis': 'A4:E4',
                                       })

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                  <dimension ref="A1:E4"/>
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
                        <v>1</v>
                      </c>
                      <c r="B4">
                        <v>2</v>
                      </c>
                      <c r="C4">
                        <v>3</v>
                      </c>
                      <c r="D4">
                        <v>4</v>
                      </c>
                      <c r="E4">
                        <v>5</v>
                      </c>
                    </row>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                        <x14:sparklineGroup manualMin="0" lineWeight="6" dateAxis="1" displayEmptyCellsAs="gap" displayHidden="1" minAxisType="custom" maxAxisType="group">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <xm:f>Sheet1!A4:E4</xm:f>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A3:E3</xm:f>
                              <xm:sqref>F3</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup lineWeight="2.25" displayEmptyCellsAs="span" minAxisType="group" maxAxisType="group">
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
                        <x14:sparklineGroup manualMax="0.5" manualMin="-0.5" lineWeight="0.25" markers="1" high="1" low="1" first="1" last="1" negative="1" displayXAxis="1" minAxisType="custom" maxAxisType="custom" rightToLeft="1">
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
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
