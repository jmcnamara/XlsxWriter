###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.format import Format
from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """

    def test_assemble_xml_file(self):
        """Test autofit of columns with user defined properties."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.str_table = SharedStringTable()

        worksheet.select()
        cell_format = Format({"xf_index": 1})

        # Autofit should widen a column with a user defined width, or set the
        # width of a column with a user defined format, without hiding it.
        worksheet.set_column(0, 0, 2)
        worksheet.set_column(2, 2, None, cell_format)
        worksheet.write_string(0, 0, "hello world")
        worksheet.write_string(0, 2, "hello world")
        worksheet.autofit()

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="A1:C1"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <cols>
                    <col min="1" max="1" width="11.140625" bestFit="1" customWidth="1"/>
                    <col min="3" max="3" width="11.140625" style="1" bestFit="1" customWidth="1"/>
                  </cols>
                  <sheetData>
                    <row r="1" spans="1:3">
                      <c r="A1" t="s">
                        <v>0</v>
                      </c>
                      <c r="C1" s="1" t="s">
                        <v>0</v>
                      </c>
                    </row>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
