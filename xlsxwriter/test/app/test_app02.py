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

from xlsxwriter.app import App

from ..helperfunctions import _xml_to_list


class TestAssembleApp(unittest.TestCase):
    """
    Test assembling a complete App file.

    """

    def test_assemble_xml_file(self):
        """Test writing an App file."""
        self.maxDiff = None

        fh = StringIO()
        app = App()
        app._set_filehandle(fh)

        app._add_part_name("Sheet1")
        app._add_part_name("Sheet2")
        app._add_heading_pair(("Worksheets", 2))

        app._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                  <Application>Microsoft Excel</Application>
                  <DocSecurity>0</DocSecurity>
                  <ScaleCrop>false</ScaleCrop>
                  <HeadingPairs>
                    <vt:vector size="2" baseType="variant">
                      <vt:variant>
                        <vt:lpstr>Worksheets</vt:lpstr>
                      </vt:variant>
                      <vt:variant>
                        <vt:i4>2</vt:i4>
                      </vt:variant>
                    </vt:vector>
                  </HeadingPairs>
                  <TitlesOfParts>
                    <vt:vector size="2" baseType="lpstr">
                      <vt:lpstr>Sheet1</vt:lpstr>
                      <vt:lpstr>Sheet2</vt:lpstr>
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
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
