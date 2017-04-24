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
        """Test writing a data table."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()

        worksheet.write('A1', 0)
        worksheet.write('B1', 0)

        worksheet.write('A3', 100)
        worksheet.write('A4', 200)

        worksheet.write('B2', 3)
        worksheet.write('C2', 4)

        worksheet.write_formula('A2', '=SUM(A1, B1)')

        worksheet.write_data_table('B3', 'C4', 'A1', 'B1')

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                    <dimension ref="A1:C4"/>
                    <sheetViews>
                        <sheetView tabSelected="1" workbookViewId="0"/>
                    </sheetViews>
                    <sheetFormatPr defaultRowHeight="15"/>
                    <sheetData>
                        <row r="1" spans="1:3">
                            <c r="A1">
                                <v>0</v>
                            </c>
                            <c r="B1">
                                <v>0</v>
                            </c>
                        </row>
                        <row r="2" spans="1:3">
                            <c r="A2">
                                <f>SUM(A1, B1)</f>
                                <v>0</v>
                            </c>
                            <c r="B2">
                                <v>3</v>
                            </c>
                            <c r="C2">
                                <v>4</v>
                            </c>
                        </row>
                        <row r="3" spans="1:3">
                            <c r="A3">
                                <v>100</v>
                            </c>
                            <c r="B3">
                                <f r="B3" t="dataTable" ref="B3:C4" ca="1" dt2D="1" dtr="1" r1="A1" r2="B1"/>
                            </c>
                        </row>
                        <row r="4" spans="1:3">
                            <c r="A4">
                                <v>200</v>
                            </c>
                        </row>
                    </sheetData>
                    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
