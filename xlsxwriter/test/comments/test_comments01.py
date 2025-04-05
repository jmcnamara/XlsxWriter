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

from ...comments import Comments, CommentType
from ..helperfunctions import _xml_to_list


class TestAssembleComments(unittest.TestCase):
    """
    Test assembling a complete Comments file.

    """

    def test_assemble_xml_file(self):
        """Test writing a comments with no cell data."""
        self.maxDiff = None

        fh = StringIO()
        comments = Comments()
        comments._set_filehandle(fh)

        comment = CommentType(1, 1, "Some text")
        comment.vertices = [2, 0, 4, 4, 143, 10, 128, 74]
        comment.author = "John"

        comments._assemble_xml_file([comment])

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <authors>
                    <author>John</author>
                  </authors>
                  <commentList>
                    <comment ref="B2" authorId="0">
                      <text>
                        <r>
                          <rPr>
                            <sz val="8"/>
                            <color indexed="81"/>
                            <rFont val="Tahoma"/>
                            <family val="2"/>
                          </rPr>
                          <t>Some text</t>
                        </r>
                      </text>
                    </comment>
                  </commentList>
                </comments>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
