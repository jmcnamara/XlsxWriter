###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ..helperfunctions import _xml_to_list
from ...chartsheet import Chartsheet


class TestAssembleChartsheet(unittest.TestCase):
    """
    Test assembling a complete Chartsheet file.

    """
    def test_assemble_xml_file(self):
        """Test writing a chartsheet with no cell data."""
        self.maxDiff = None

        fh = StringIO()
        chartsheet = Chartsheet()
        chartsheet._set_filehandle(fh)

        chartsheet.drawing = 1

        chartsheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheetPr/>
                  <sheetViews>
                    <sheetView workbookViewId="0"/>
                  </sheetViews>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                  <drawing r:id="rId1"/>
                </chartsheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
