###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list


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

        worksheet.conditional_format(
            "A1",
            {"type": "cell", "criteria": "==", "value": '"Test A2"', "format": None},
        )

        worksheet._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <dimension ref="A1"/>
                <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15"/>
                <sheetData/>
                <conditionalFormatting sqref="A1">
                    <cfRule type="cellIs" priority="1" operator="equal">
                    <formula>"Test A2"</formula>
                    </cfRule>
                </conditionalFormatting>
                <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
            """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
