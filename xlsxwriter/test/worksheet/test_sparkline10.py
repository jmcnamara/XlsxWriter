###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
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

        # Set up sparklines.
        worksheet.add_sparkline('F1', {'range': 'A1:E1',
                                       'high_point': True,
                                       'low_point': True,
                                       'negative_points': True,
                                       'first_point': True,
                                       'last_point': True,
                                       'markers': True,
                                       'series_color': '#C00000',
                                       'negative_color': '#FF0000',
                                       'markers_color': '#FFC000',
                                       'first_color': '#00B050',
                                       'last_color': '#00B0F0',
                                       'high_color': '#FFFF00',
                                       'low_color': '#92D050',
                                       })

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                  <dimension ref="A1:E1"/>
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
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                        <x14:sparklineGroup displayEmptyCellsAs="gap" markers="1" high="1" low="1" first="1" last="1" negative="1">
                          <x14:colorSeries rgb="FFC00000"/>
                          <x14:colorNegative rgb="FFFF0000"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FFFFC000"/>
                          <x14:colorFirst rgb="FF00B050"/>
                          <x14:colorLast rgb="FF00B0F0"/>
                          <x14:colorHigh rgb="FFFFFF00"/>
                          <x14:colorLow rgb="FF92D050"/>
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
