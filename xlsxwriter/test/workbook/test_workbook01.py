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

from xlsxwriter.workbook import Workbook

from ..helperfunctions import _xml_to_list


class TestAssembleWorkbook(unittest.TestCase):
    """
    Test assembling a complete Workbook file.

    """

    def test_assemble_xml_file(self):
        """Test writing a workbook with 1 worksheet."""
        self.maxDiff = None

        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        workbook.add_worksheet()

        workbook._assemble_xml_file()
        workbook.fileclosed = 1

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
                  <workbookPr defaultThemeVersion="124226"/>
                  <bookViews>
                    <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
                  </bookViews>
                  <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                  </sheets>
                  <calcPr calcId="124519" fullCalcOnLoad="1"/>
                </workbook>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
