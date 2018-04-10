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
from ...format import Format


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """
    def test_assemble_xml_file(self):
        """Test writing a worksheet with row formatting set."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        cell_format = Format({'xf_index': 1})

        worksheet.set_row(1, 30)
        worksheet.set_row(3, None, None, {'hidden': 1})
        worksheet.set_row(6, None, cell_format)
        worksheet.set_row(9, 3)
        worksheet.set_row(12, 24, None, {'hidden': 1})
        worksheet.set_row(14, 0)

        worksheet.select()
        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="A2:A15"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <sheetData>
                    <row r="2" ht="30" customHeight="1"/>
                    <row r="4" hidden="1"/>
                    <row r="7" s="1" customFormat="1"/>
                    <row r="10" ht="3" customHeight="1"/>
                    <row r="13" ht="24" hidden="1" customHeight="1"/>
                    <row r="15" hidden="1"/>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
