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
from xlsxwriter.table import Table
from xlsxwriter.workbook import WorksheetMeta
from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list

EXPECTED_TABLE_XML = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C2:F14" totalsRowCount="1">
      <autoFilter ref="C2:F13"/>
      <tableColumns count="4">
        <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
        <tableColumn id="2" name="Column2"/>
        <tableColumn id="3" name="Column3"/>
        <tableColumn id="4" name="Column4" totalsRowFunction="count" dataDxfId="0">
          <calculatedColumnFormula>SUM(Table1[[#This Row],[Column1]:[Column3]])</calculatedColumnFormula>
        </tableColumn>
      </tableColumns>
      <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
    </table>
"""


class TestAssembleTable(unittest.TestCase):
    """
    Test assembling a complete Table file.

    """

    maxDiff = None

    def setUp(self):
        self.worksheet = Worksheet()
        self.worksheet.worksheet_meta = WorksheetMeta()
        self.worksheet.str_table = SharedStringTable()

        self.fh = StringIO()
        self.table = Table()
        self.table._set_filehandle(self.fh)

    def _add_table_with_total_row_and_formula(self):
        dxf_format = Format()
        dxf_format.dxf_index = 0

        # Set the table properties.
        self.worksheet.add_table(
            "C2:F14",
            {
                "total_row": 1,
                "columns": [
                    {"total_string": "Total"},
                    {},
                    {},
                    {
                        "total_function": "count",
                        "format": dxf_format,
                        "formula": "SUM(Table1[[#This Row],[Column1]:[Column3]])",
                    },
                ],
            },
        )
        self.worksheet._prepare_tables(1, {})

    def _assemble_and_get_xml(self):
        self.table._set_properties(self.worksheet.tables[0])
        self.table._assemble_xml_file()
        return _xml_to_list(self.fh.getvalue())

    def test_assemble_xml_file(self):
        """Test writing a table"""
        self._add_table_with_total_row_and_formula()

        exp = _xml_to_list(EXPECTED_TABLE_XML)
        got = self._assemble_and_get_xml()

        self.assertEqual(exp, got)
