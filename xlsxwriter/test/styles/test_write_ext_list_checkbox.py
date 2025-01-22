###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from ...styles import Styles
from ..helperfunctions import _xml_to_list


class TestWriteExtLstCheckbox(unittest.TestCase):
    """
    Test the Styles _write_ext_lst_checkbox() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_ext_lst_checkbox(self):
        """Test the _write_ext_lst_checkbox() method"""

        self.styles._write_ext_lst_checkbox()

        exp = _xml_to_list(
            """
                <extLst>
                    <ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag">
                        <xfpb:xfComplement i="0"/>
                    </ext>
                </extLst>
            """
        )
        got = _xml_to_list(self.fh.getvalue())

        self.assertEqual(got, exp)
