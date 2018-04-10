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
        """Test writing a worksheet with column formatting set."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        cell_format = Format({'xf_index': 1})

        worksheet.set_column(1, 3, 5)
        worksheet.set_column(5, 5, 8, None, {'hidden': True})
        worksheet.set_column(7, 7, None, cell_format)
        worksheet.set_column(9, 9, 2)
        worksheet.set_column(11, 11, None, None, {'hidden': True})

        worksheet.select()
        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="F1:H1"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <cols>
                    <col min="2" max="4" width="5.7109375" customWidth="1"/>
                    <col min="6" max="6" width="8.7109375" hidden="1" customWidth="1"/>
                    <col min="8" max="8" width="9.140625" style="1"/>
                    <col min="10" max="10" width="2.7109375" customWidth="1"/>
                    <col min="12" max="12" width="0" hidden="1" customWidth="1"/>
                  </cols>
                  <sheetData/>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)

    def test_assemble_xml_file_A1(self):
        """
        Test writing a worksheet with column formatting set using
        A1 Notation.
        """
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        cell_format = Format({'xf_index': 1})

        worksheet.set_column('B:D', 5)
        worksheet.set_column('F:F', 8, None, {'hidden': True})
        worksheet.set_column('H:H', None, cell_format)
        worksheet.set_column('J:J', 2)
        worksheet.set_column('L:L', None, None, {'hidden': True})

        worksheet.select()
        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="F1:H1"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <cols>
                    <col min="2" max="4" width="5.7109375" customWidth="1"/>
                    <col min="6" max="6" width="8.7109375" hidden="1" customWidth="1"/>
                    <col min="8" max="8" width="9.140625" style="1"/>
                    <col min="10" max="10" width="2.7109375" customWidth="1"/>
                    <col min="12" max="12" width="0" hidden="1" customWidth="1"/>
                  </cols>
                  <sheetData/>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
