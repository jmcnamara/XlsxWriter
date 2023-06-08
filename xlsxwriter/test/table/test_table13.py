###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ..helperfunctions import _xml_to_list
from ...table import Table
from ...worksheet import Worksheet
from ...workbook import WorksheetMeta
from ...sharedstrings import SharedStringTable


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

        # Set the table properties.
        worksheet.add_table(
            "B2:D8",
            {
                "total_row": 1,
                "columns": [
                    {"total_string": "Total"},
                    {},
                    {
                        "total_formula": "=SUBTOTAL(9, [Column1]) + SUBTOTAL(9, [Column3])",
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
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="B2:D8" totalsRowCount="1">
                  <autoFilter ref="B2:D7"/>
                  <tableColumns count="3">
                    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3" totalsRowFunction="custom">
                      <totalsRowFormula>SUBTOTAL(9, [Column1]) + SUBTOTAL(9, [Column3])</totalsRowFormula>
                    </tableColumn>
                  </tableColumns>
                  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
