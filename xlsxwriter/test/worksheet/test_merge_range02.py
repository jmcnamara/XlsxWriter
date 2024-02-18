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
from ...format import Format
from ...sharedstrings import SharedStringTable


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """

    def test_assemble_xml_file(self):
        """Test merged cell range"""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.str_table = SharedStringTable()
        worksheet.select()
        cell_format1 = Format({"xf_index": 1})
        cell_format2 = Format({"xf_index": 2})

        worksheet.merge_range("B3:C3", "Foo", cell_format1)
        worksheet.merge_range("A2:D2", "", cell_format2)

        worksheet.select()
        worksheet._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="A2:D3"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <sheetData>
                    <row r="2" spans="1:4">
                      <c r="A2" s="2"/>
                      <c r="B2" s="2"/>
                      <c r="C2" s="2"/>
                      <c r="D2" s="2"/>
                    </row>
                    <row r="3" spans="1:4">
                      <c r="B3" s="1" t="s">
                        <v>0</v>
                      </c>
                      <c r="C3" s="1"/>
                    </row>
                  </sheetData>
                  <mergeCells count="2">
                    <mergeCell ref="B3:C3"/>
                    <mergeCell ref="A2:D2"/>
                  </mergeCells>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)

    def test_assemble_xml_file_write(self):
        """Test writing a worksheet with a blank cell with write() method."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        cell_format = Format({"xf_index": 1})

        # No format. Should be ignored.
        worksheet.write(0, 0, None)

        worksheet.write(1, 2, None, cell_format)

        worksheet.select()
        worksheet._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="C2"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <sheetData>
                    <row r="2" spans="3:3">
                      <c r="C2" s="1"/>
                    </row>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)

    def test_assemble_xml_file_A1(self):
        """Test writing a worksheet with a blank cell with A1 notation."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        cell_format = Format({"xf_index": 1})

        # No format. Should be ignored.
        worksheet.write_blank("A1", None)

        worksheet.write_blank("C2", None, cell_format)

        worksheet.select()
        worksheet._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="C2"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <sheetData>
                    <row r="2" spans="3:3">
                      <c r="C2" s="1"/>
                    </row>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
