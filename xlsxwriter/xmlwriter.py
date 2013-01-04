###############################################################################
#
# XMLwriter - A base class for XlsxWriter classes.
#
# Used in conjunction with XlsxWriter.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import re

class XMLwriter:
    """Simple XML writer class."""


    def __init__(self, filehandle=None):
        self.fh = filehandle
        self.escapes = re.compile('["&<>]')


    def xml_declaration(self):
        """Write the XML declaration."""

        self.fh.write(
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n""")


    def xml_start_tag(self, tag, attributes=[]):
        """Write an XML start tag with optional attributes."""

        while attributes:
            key   = attributes.pop(0)
            value = attributes.pop(0)
            value = self._escape_attributes(value)
            tag   = tag + ' %s="%s"' % (key, value)

        self.fh.write("<%s>" % tag)


    def xml_end_tag(self, tag):
        """Write an XML end tag."""

        self.fh.write("</%s>" % tag)


    def xml_empty_tag(self, tag, attributes=[]):
        """Write an empty XML tag with optional attributes."""

        while attributes:
            key   = attributes.pop(0)
            value = attributes.pop(0)
            value = self._escape_attributes(value)
            tag   = tag + ' %s="%s"' % (key, value)

        self.fh.write("<%s/>" % tag)


    def xml_data_element(self, tag, data, attributes=[]):
        """Write an XML element containing data with optional attributes."""

        end_tag = tag

        while attributes:
            key   = attributes.pop(0)
            value = attributes.pop(0)
            value = self._escape_attributes(value)
            tag   = tag + ' %s="%s"' % (key, value)

        data = self._escape_data(data)

        self.fh.write("<%s>%s</%s>" % (tag, data, end_tag))


    def _escape_attributes(self, str):
        """Escape XML characters in attributes."""

        if self.escapes.match(str) == None:
            return str

        str = str.replace('&', '&amp;')
        str = str.replace('"', '&quot;')
        str = str.replace('<', '&lt;')
        str = str.replace('>', '&gt;')

        return str


    def _escape_data(self, str):
        """Escape XML characters in attributes."""

        if self.escapes.match(str) == None:
            return str

        str = str.replace('&', '&amp;')
        str = str.replace('<', '&lt;')
        str = str.replace('>', '&gt;')

        return str
