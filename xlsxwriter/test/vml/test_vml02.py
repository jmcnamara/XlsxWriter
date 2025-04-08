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

from ...vml import ButtonType, Vml
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

        button = ButtonType(1, 2, 17, 64, 1)
        button.vertices = [2, 1, 0, 0, 3, 2, 0, 0, 128, 20, 64, 20]

        vml._assemble_xml_file(
            1,
            1024,
            None,
            [button],
        )

        exp = _vml_to_list(
            """
                <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                  <o:shapelayout v:ext="edit">
                    <o:idmap v:ext="edit" data="1"/>
                  </o:shapelayout>
                  <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">
                    <v:stroke joinstyle="miter"/>
                    <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
                    <o:lock v:ext="edit" shapetype="t"/>
                  </v:shapetype>
                  <v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute;margin-left:96pt;margin-top:15pt;width:48pt;height:15pt;z-index:1;mso-wrap-style:tight" o:button="t" fillcolor="buttonFace [67]" strokecolor="windowText [64]" o:insetmode="auto">
                    <v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>
                    <o:lock v:ext="edit" rotation="t"/>
                    <v:textbox style="mso-direction-alt:auto" o:singleclick="f">
                      <div style="text-align:center">
                        <font face="Calibri" size="220" color="#000000">Button 1</font>
                      </div>
                    </v:textbox>
                    <x:ClientData ObjectType="Button">
                      <x:Anchor>2, 0, 1, 0, 3, 0, 2, 0</x:Anchor>
                      <x:PrintObject>False</x:PrintObject>
                      <x:AutoFill>False</x:AutoFill>
                      <x:FmlaMacro>[0]!Button1_Click</x:FmlaMacro>
                      <x:TextHAlign>Center</x:TextHAlign>
                      <x:TextVAlign>Center</x:TextVAlign>
                    </x:ClientData>
                  </v:shape>
                </xml>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
