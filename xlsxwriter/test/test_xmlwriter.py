###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from StringIO  import StringIO
from ..xmlwriter import XMLwriter

class TestXMLwriter(unittest.TestCase):
    """Test the XML Writer class."""


    def setUp(self):
        self.fh     = StringIO()
        self.writer = XMLwriter(self.fh)


    def test_xml_declaration(self):
        """Test xml_declaration()"""

        self.writer.xml_declaration()

        expected = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_start_tag(self):
        """Test xml_start_tag() with no attributes"""

        self.writer.xml_start_tag('foo')

        expected = """<foo>"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_start_tag_with_attributes(self):
        """Test xml_start_tag() with attributes"""

        self.writer.xml_start_tag('foo', ['span', '8', 'baz', '7'])

        expected = """<foo span="8" baz="7">"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_start_tag_with_attributes_to_escape(self):
        """Test xml_start_tag() with attributes requiring escaping"""

        self.writer.xml_start_tag('foo', ['span', '&<>"'])

        expected = """<foo span="&amp;&lt;&gt;&quot;">"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_end_tag(self):
        """Test xml_end_tag()"""

        self.writer.xml_end_tag('foo')

        expected = """</foo>"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_empty_tag(self):
        """Test xml_empty_tag()"""

        self.writer.xml_empty_tag('foo')

        expected = """<foo/>"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_empty_tag_with_attributes(self):
        """Test xml_empty_tag() with attributes"""

        self.writer.xml_empty_tag('foo', ['span', '8'])

        expected = """<foo span="8"/>"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_data_element(self):
        """Test xml_data_element()"""

        self.writer.xml_data_element('foo', 'bar')

        expected = """<foo>bar</foo>"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_data_element_with_attributes(self):
        """Test xml_data_element() with attributes"""

        self.writer.xml_data_element('foo', 'bar', ['span', '8'])

        expected = """<foo span="8">bar</foo>"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)


    def test_xml_data_element_with_escapes(self):
        """Test xml_data_element() with data requiring escaping"""

        self.writer.xml_data_element('foo', '&<>"', ['span', '8'])

        expected = """<foo span="8">&amp;&lt;&gt;"</foo>"""
        got      = self.fh.getvalue()

        self.assertEqual(got, expected)



if __name__ == '__main__':
    unittest.main()
