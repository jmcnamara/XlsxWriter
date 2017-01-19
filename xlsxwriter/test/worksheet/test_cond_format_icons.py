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
        """Test writing a worksheet with conditional formatting."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()

        worksheet.write('A1', 1234)
        worksheet.write('A2', 0)
        worksheet.write('A3', -152)
        worksheet.write('A4', 1)
        worksheet.write('A5', 0)
        worksheet.write('A6', -1)
        worksheet.write('A7', 12)
        worksheet.write('A8', 0)
        worksheet.write('A9', -1)
        worksheet.write('A10', 100)
        worksheet.write('A11', 0)
        worksheet.write('A12', -145)
        worksheet.write('A13', 199206)
        worksheet.write('A14', 0)
        worksheet.write('A15', -100000)
        worksheet.write('A16', 1234)
        worksheet.write('A17', 0)
        worksheet.write('A18', -1)

        worksheet.conditional_format('A1:A17',
                                     {'type': 'icon_set',
                                      'icon_type': '3Arrows',
                                      })

        worksheet.conditional_format('A1:A17',
                                     {'type': 'icon_set',
                                      'icon_type': '3Arrows',
                                      'show_value': 0,
                                      })

        worksheet.conditional_format('A1:A17',
                                     {'type': 'icon_set',
                                      'icon_type': '3Arrows',
                                      'show_value': 0,
                                      'min_value': 0,
                                      'mid_value': 0,
                                      'max_value': 0,
                                      'min_type': 'percent',
                                      'mid_type': 'num',
                                      'max_type': 'num',
                                      'gte': 0,
                                      })

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                    <dimension ref="A1:A18"/>
                    <sheetViews>
                        <sheetView tabSelected="1" workbookViewId="0"/>
                    </sheetViews>
                    <sheetFormatPr defaultRowHeight="15"/>
                    <sheetData>
                        <row r="1" spans="1:1">
                            <c r="A1">
                                <v>1234</v>
                            </c>
                        </row>
                            <row r="2" spans="1:1">
                                <c r="A2">
                                    <v>0</v>
                                </c>
                            </row>
                            <row r="3" spans="1:1">
                                <c r="A3">
                                    <v>-152</v>
                                </c>
                            </row>
                            <row r="4" spans="1:1">
                                <c r="A4">
                                    <v>1</v>
                                </c>
                            </row>
                                <row r="5" spans="1:1">
                                    <c r="A5">
                                        <v>0</v>
                                    </c>
                                </row>
                                    <row r="6" spans="1:1">
                                        <c r="A6">
                                            <v>-1</v>
                                        </c>
                                    </row>
                                    <row r="7" spans="1:1">
                                        <c r="A7">
                                            <v>12</v>
                                        </c>
                                    </row>
                                    <row r="8" spans="1:1">
                                        <c r="A8">
                                            <v>0</v>
                                        </c>
                                    </row>
                                    <row r="9" spans="1:1">
                                        <c r="A9">
                                            <v>-1</v>
                                        </c>
                                    </row>
                                    <row r="10" spans="1:1">
                                        <c r="A10">
                                            <v>100</v>
                                        </c>
                                    </row>
                                    <row r="11" spans="1:1">
                                        <c r="A11">
                                            <v>0</v>
                                        </c>
                                    </row>
                                    <row r="12" spans="1:1">
                                        <c r="A12">
                                            <v>-145</v>
                                        </c>
                                    </row>
                                    <row r="13" spans="1:1">
                                        <c r="A13">
                                            <v>199206</v>
                                        </c>
                                    </row>
                                    <row r="14" spans="1:1">
                                        <c r="A14">
                                            <v>0</v>
                                        </c>
                                    </row>
                                    <row r="15" spans="1:1">
                                        <c r="A15">
                                            <v>-100000</v>
                                        </c>
                                    </row>
                                    <row r="16" spans="1:1">
                                        <c r="A16">
                                            <v>1234</v>
                                        </c>
                                    </row>
                                    <row r="17" spans="1:1">
                                        <c r="A17">
                                            <v>0</v>
                                        </c>
                                    </row>
                                    <row r="18" spans="1:1">
                                        <c r="A18">
                                            <v>-1</v>
                                        </c>
                                    </row>
                                </sheetData>
                                <conditionalFormatting sqref="A1:A17">
                                    <cfRule type="iconSet" priority="1">
                                        <iconSet iconSet="3Arrows" showValue="1">
                                            <cfvo type="percent" val="0"/>
                                            <cfvo type="percent" val="33"/>
                                            <cfvo type="percent" val="67"/>
                                        </iconSet>
                                    </cfRule>
                                    <cfRule type="iconSet" priority="2">
                                        <iconSet iconSet="3Arrows" showValue="0">
                                            <cfvo type="percent" val="0"/>
                                            <cfvo type="percent" val="33"/>
                                            <cfvo type="percent" val="67"/>
                                        </iconSet>
                                    </cfRule>
                                    <cfRule type="iconSet" priority="3">
                                        <iconSet iconSet="3Arrows" showValue="0">
                                            <cfvo type="percent" val="0"/>
                                            <cfvo type="num" val="0"/>
                                            <cfvo type="num" val="0" gte="0"/>
                                        </iconSet>
                                    </cfRule>
                                    </conditionalFormatting>
                                    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                                </worksheet>
        """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
