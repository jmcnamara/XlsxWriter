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

from xlsxwriter.image import Image
from xlsxwriter.vml import Vml

from ..helperfunctions import _vml_to_list, _xml_to_list


class TestAssembleVml(unittest.TestCase):
    """
    Test assembling a complete Vml file.

    """

    def test_assemble_xml_file(self):
        """Test writing a vml with no cell data."""
        self.maxDiff = None

        fh = StringIO()
        vml = Vml()
        vml._set_filehandle(fh)

        image = Image("xlsxwriter/test/comparison/images/red.png")
        image._header_position = "CH"
        image._ref_id = 1
        image.image_name = "red"

        vml._assemble_xml_file(1, 1024, None, None, [image])

        exp = _vml_to_list(
            """
                <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                  <o:shapelayout v:ext="edit">
                    <o:idmap v:ext="edit" data="1"/>
                  </o:shapelayout>
                  <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
                    <v:stroke joinstyle="miter"/>
                    <v:formulas>
                      <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                      <v:f eqn="sum @0 1 0"/>
                      <v:f eqn="sum 0 0 @1"/>
                      <v:f eqn="prod @2 1 2"/>
                      <v:f eqn="prod @3 21600 pixelWidth"/>
                      <v:f eqn="prod @3 21600 pixelHeight"/>
                      <v:f eqn="sum @0 0 1"/>
                      <v:f eqn="prod @6 1 2"/>
                      <v:f eqn="prod @7 21600 pixelWidth"/>
                      <v:f eqn="sum @8 21600 0"/>
                      <v:f eqn="prod @7 21600 pixelHeight"/>
                      <v:f eqn="sum @10 21600 0"/>
                    </v:formulas>
                    <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
                    <o:lock v:ext="edit" aspectratio="t"/>
                  </v:shapetype>
                  <v:shape id="CH" o:spid="_x0000_s1025" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:24pt;height:24pt;z-index:1">
                    <v:imagedata o:relid="rId1" o:title="red"/>
                    <o:lock v:ext="edit" rotation="t"/>
                  </v:shape>
                </xml>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
