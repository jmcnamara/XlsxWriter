###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...xmlwriter import XMLwriter


class TestXMLwriter(unittest.TestCase):
    """
    Test the XML Writer class.

    """

    def setUp(self):
        self.fh = StringIO()
        self.writer = XMLwriter()
        self.writer._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test _xml_declaration()"""

        self.writer._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_start_tag(self):
        """Test _xml_start_tag() with no attributes"""

        self.writer._xml_start_tag("foo")

        exp = """<foo>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_start_tag_with_attributes(self):
        """Test _xml_start_tag() with attributes"""

        self.writer._xml_start_tag("foo", [("span", "8"), ("baz", "7")])

        exp = """<foo span="8" baz="7">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_start_tag_with_attributes_to_escape(self):
        """Test _xml_start_tag() with attributes requiring escaping"""

        self.writer._xml_start_tag("foo", [("span", '&<>"')])

        exp = """<foo span="&amp;&lt;&gt;&quot;">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_start_tag_unencoded(self):
        """Test _xml_start_tag_unencoded() with attributes"""

        self.writer._xml_start_tag_unencoded("foo", [("span", '&<>"')])

        exp = """<foo span="&<>"">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_end_tag(self):
        """Test _xml_end_tag()"""

        self.writer._xml_end_tag("foo")

        exp = """</foo>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_empty_tag(self):
        """Test _xml_empty_tag()"""

        self.writer._xml_empty_tag("foo")

        exp = """<foo/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_empty_tag_with_attributes(self):
        """Test _xml_empty_tag() with attributes"""

        self.writer._xml_empty_tag("foo", [("span", "8")])

        exp = """<foo span="8"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_empty_tag_unencoded(self):
        """Test _xml_empty_tag_unencoded() with attributes"""

        self.writer._xml_empty_tag_unencoded("foo", [("span", "&")])

        exp = """<foo span="&"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_data_element(self):
        """Test _xml_data_element()"""

        self.writer._xml_data_element("foo", "bar")

        exp = """<foo>bar</foo>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_data_element_with_attributes(self):
        """Test _xml_data_element() with attributes"""

        self.writer._xml_data_element("foo", "bar", [("span", "8")])

        exp = """<foo span="8">bar</foo>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_data_element_with_escapes(self):
        """Test _xml_data_element() with data requiring escaping"""

        self.writer._xml_data_element("foo", '&<>"', [("span", "8")])

        exp = """<foo span="8">&amp;&lt;&gt;"</foo>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_string_element(self):
        """Test _xml_string_element()"""

        self.writer._xml_string_element(99, [("span", "8")])

        exp = """<c span="8" t=\"s\"><v>99</v></c>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_si_element(self):
        """Test _xml_si_element()"""

        self.writer._xml_si_element("foo", [("span", "8")])

        exp = """<si><t span="8">foo</t></si>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_rich_si_element(self):
        """Test _xml_rich_si_element()"""

        self.writer._xml_rich_si_element("foo")

        exp = """<si>foo</si>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_number_element(self):
        """Test _xml_number_element()"""

        self.writer._xml_number_element(99, [("span", "8")])

        exp = """<c span="8"><v>99</v></c>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_xml_formula_element(self):
        """Test _xml_formula_element()"""

        self.writer._xml_formula_element("1+2", 3, [("span", "8")])

        exp = """<c span="8"><f>1+2</f><v>3</v></c>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
