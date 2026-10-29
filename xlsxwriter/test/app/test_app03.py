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

from xlsxwriter.app import App

from ..helperfunctions import _xml_to_list

EXPECTED_APP_XML_PRINT_TITLES = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <Application>Microsoft Excel</Application>
      <DocSecurity>0</DocSecurity>
      <ScaleCrop>false</ScaleCrop>
      <HeadingPairs>
        <vt:vector size="4" baseType="variant">
          <vt:variant>
            <vt:lpstr>Worksheets</vt:lpstr>
          </vt:variant>
          <vt:variant>
            <vt:i4>1</vt:i4>
          </vt:variant>
          <vt:variant>
            <vt:lpstr>Named Ranges</vt:lpstr>
          </vt:variant>
          <vt:variant>
            <vt:i4>1</vt:i4>
          </vt:variant>
        </vt:vector>
      </HeadingPairs>
      <TitlesOfParts>
        <vt:vector size="2" baseType="lpstr">
          <vt:lpstr>Sheet1</vt:lpstr>
          <vt:lpstr>Sheet1!Print_Titles</vt:lpstr>
        </vt:vector>
      </TitlesOfParts>
      <Company>
      </Company>
      <LinksUpToDate>false</LinksUpToDate>
      <SharedDoc>false</SharedDoc>
      <HyperlinksChanged>false</HyperlinksChanged>
      <AppVersion>12.0000</AppVersion>
    </Properties>
"""


class TestAssembleApp(unittest.TestCase):
    """
    Test assembling a complete App file.

    """

    maxDiff = None

    def setUp(self):
        self.fh = StringIO()
        self.app = App()
        self.app._set_filehandle(self.fh)

    def _assemble_and_get_xml(self):
        self.app._assemble_xml_file()
        return _xml_to_list(self.fh.getvalue())

    def test_worksheet_and_named_range_produces_expected_app_xml(self):
        self.app._add_part_name("Sheet1")
        self.app._add_part_name("Sheet1!Print_Titles")
        self.app._add_heading_pair(("Worksheets", 1))
        self.app._add_heading_pair(("Named Ranges", 1))

        expected = _xml_to_list(EXPECTED_APP_XML_PRINT_TITLES)

        self.assertEqual(expected, self._assemble_and_get_xml())
