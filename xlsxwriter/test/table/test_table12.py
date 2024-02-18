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
from ...table import Table
from ...worksheet import Worksheet
from ...workbook import WorksheetMeta
from ...sharedstrings import SharedStringTable
from ...format import Format


class TestAssembleTable(unittest.TestCase):
    """
    Test assembling a complete Table file.

    """

    def test_assemble_xml_file(self):
        """Test writing a table"""
        self.maxDiff = None

        worksheet = Worksheet()
        worksheet.worksheet_meta = WorksheetMeta()
        worksheet.str_table = SharedStringTable()
        dxf_format = Format()
        dxf_format.dxf_index = 0

        # Set the table properties.
        worksheet.add_table(
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
                        "formula": "=SUM(Table1[@[Column1]:[Column3]])",
                    },
                ],
            },
        )
        worksheet._prepare_tables(1, {})

        fh = StringIO()
        table = Table()
        table._set_filehandle(fh)

        table._set_properties(worksheet.tables[0])

        table._assemble_xml_file()

        exp = _xml_to_list(
            """
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
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
