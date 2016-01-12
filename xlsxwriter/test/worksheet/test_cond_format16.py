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

        worksheet.write('A1', 1)
        worksheet.write('A2', 2)
        worksheet.write('A3', 3)
        worksheet.write('A4', 4)
        worksheet.write('A5', 5)
        worksheet.write('A6', 6)
        worksheet.write('A7', 7)
        worksheet.write('A8', 8)
        worksheet.write('A9', 9)
        worksheet.write('A10', 10)
        worksheet.write('A11', 11)
        worksheet.write('A12', 12)

        worksheet.conditional_format('A1:A12',
                                     {'type': '3_color_scale',
                                      'min_color': "#C5D9F1",
                                      'mid_color': "#8DB4E3",
                                      'max_color': "#538ED5",
                                      })

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
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
                    <row r="10" spans="1:1">
                      <c r="A10">
                        <v>10</v>
                      </c>
                    </row>
                    <row r="11" spans="1:1">
                      <c r="A11">
                        <v>11</v>
                      </c>
                    </row>
                    <row r="12" spans="1:1">
                      <c r="A12">
                        <v>12</v>
                      </c>
                    </row>
                  </sheetData>
                  <conditionalFormatting sqref="A1:A12">
                    <cfRule type="colorScale" priority="1">
                      <colorScale>
                        <cfvo type="min" val="0"/>
                        <cfvo type="percentile" val="50"/>
                        <cfvo type="max" val="0"/>
                        <color rgb="FFC5D9F1"/>
                        <color rgb="FF8DB4E3"/>
                        <color rgb="FF538ED5"/>
                      </colorScale>
                    </cfRule>
                  </conditionalFormatting>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
