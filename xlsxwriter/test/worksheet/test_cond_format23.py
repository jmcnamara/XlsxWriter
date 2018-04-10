###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
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
        """Test writing a worksheet with conditional formatting."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()

        worksheet.write('A1', 1)
        worksheet.write('A2', 2)
        worksheet.write('A3', 3)
        worksheet.write('A4', 4)
        worksheet.write('A5', 5)
        worksheet.write('A6', 6)
        worksheet.write('A7', 7)
        worksheet.write('A8', 8)

        worksheet.conditional_format('A1',
                                     {'type': 'icon_set',
                                      'icon_style': '3_arrows_gray'})

        worksheet.conditional_format('A2',
                                     {'type': 'icon_set',
                                      'icon_style': '3_traffic_lights'})

        worksheet.conditional_format('A3',
                                     {'type': 'icon_set',
                                      'icon_style': '3_signs'})

        worksheet.conditional_format('A4',
                                     {'type': 'icon_set',
                                      'icon_style': '3_symbols'})

        worksheet.conditional_format('A5',
                                     {'type': 'icon_set',
                                      'icon_style': '4_arrows_gray'})

        worksheet.conditional_format('A6',
                                     {'type': 'icon_set',
                                      'icon_style': '4_ratings'})

        worksheet.conditional_format('A7',
                                     {'type': 'icon_set',
                                      'icon_style': '5_arrows'})

        worksheet.conditional_format('A8',
                                     {'type': 'icon_set',
                                      'icon_style': '5_ratings'})

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="A1:A8"/>
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
                  </sheetData>
                  <conditionalFormatting sqref="A1">
                    <cfRule type="iconSet" priority="1">
                      <iconSet iconSet="3ArrowsGray">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="33"/>
                        <cfvo type="percent" val="67"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A2">
                    <cfRule type="iconSet" priority="2">
                      <iconSet>
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="33"/>
                        <cfvo type="percent" val="67"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A3">
                    <cfRule type="iconSet" priority="3">
                      <iconSet iconSet="3Signs">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="33"/>
                        <cfvo type="percent" val="67"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A4">
                    <cfRule type="iconSet" priority="4">
                      <iconSet iconSet="3Symbols2">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="33"/>
                        <cfvo type="percent" val="67"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A5">
                    <cfRule type="iconSet" priority="5">
                      <iconSet iconSet="4ArrowsGray">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="25"/>
                        <cfvo type="percent" val="50"/>
                        <cfvo type="percent" val="75"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A6">
                    <cfRule type="iconSet" priority="6">
                      <iconSet iconSet="4Rating">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="25"/>
                        <cfvo type="percent" val="50"/>
                        <cfvo type="percent" val="75"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A7">
                    <cfRule type="iconSet" priority="7">
                      <iconSet iconSet="5Arrows">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="20"/>
                        <cfvo type="percent" val="40"/>
                        <cfvo type="percent" val="60"/>
                        <cfvo type="percent" val="80"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A8">
                    <cfRule type="iconSet" priority="8">
                      <iconSet iconSet="5Rating">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="20"/>
                        <cfvo type="percent" val="40"/>
                        <cfvo type="percent" val="60"/>
                        <cfvo type="percent" val="80"/>
                      </iconSet>
                    </cfRule>
                  </conditionalFormatting>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
