###############################################################################
#
# XMLwriter - A base class for XlsxWriter classes.
#
# Used in conjunction with XlsxWriter.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import re


class XMLwriter(object):
    """
    Simple XML writer class.

    """

    def __init__(self):
        self.fh = None
        self.escapes = re.compile('["&<>]')

    def _set_filehandle(self, filehandle):
        # Set the writer filehandle directly. Mainly for testing.
        self.fh = filehandle

    def _xml_declaration(self):
        # Write the XML declaration.

        self.fh.write(
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n""")

    def _xml_start_tag(self, tag, attributes=[]):
        # Write an XML start tag with optional attributes.

        for key, value in attributes:
            value = self._escape_attributes(value)
            tag = tag + ' %s="%s"' % (key, value)

        self.fh.write("<%s>" % tag)

    def _xml_start_tag_unencoded(self, tag, attributes=[]):
        # Write an XML start tag with optional, unencoded, attributes.
        # This is a minor speed optimisation for elements that don't
        # need encoding.

        for key, value in attributes:
            tag = tag + ' %s="%s"' % (key, value)

        self.fh.write("<%s>" % tag)

    def _xml_end_tag(self, tag):
        # Write an XML end tag.

        self.fh.write("</%s>" % tag)

    def _xml_empty_tag(self, tag, attributes=[]):
        # Write an empty XML tag with optional attributes.

        for key, value in attributes:
            value = self._escape_attributes(value)
            tag = tag + ' %s="%s"' % (key, value)

        self.fh.write("<%s/>" % tag)

    def _xml_empty_tag_unencoded(self, tag, attributes=[]):
        # Write an XML start tag with optional, unencoded, attributes.
        # This is a minor speed optimisation for elements that don't
        # need encoding.
        for key, value in attributes:
            tag = tag + ' %s="%s"' % (key, value)

        self.fh.write("<%s/>" % tag)

    def _xml_data_element(self, tag, data, attributes=[]):
        # Write an XML element containing data with optional attributes.

        end_tag = tag

        for key, value in attributes:
            value = self._escape_attributes(value)
            tag = tag + ' %s="%s"' % (key, value)

        data = self._escape_data(data)

        self.fh.write("<%s>%s</%s>" % (tag, data, end_tag))

    def _xml_string_element(self, index, attributes=[]):
        # Optimised tag writer for <c> cell string elements in the inner loop.

        attr = ''

        for key, value in attributes:
            value = self._escape_attributes(value)
            attr = attr + ' %s="%s"' % (key, value)

        self.fh.write("""<c%s t="s"><v>%d</v></c>""" % (attr, index))

    def _xml_si_element(self, string, attributes=[]):
        # Optimised tag writer for shared strings <si> elements.

        attr = ''

        for key, value in attributes:
            value = self._escape_attributes(value)
            attr = attr + ' %s="%s"' % (key, value)

        self.fh.write("""<si><t%s>%s</t></si>""" % (attr, string))

    def _xml_rich_si_element(self, string):
        # Optimised tag writer for shared strings <si> rich string elements.

        self.fh.write("""<si>%s</si>""" %  string)

    def _xml_number_element(self, number, attributes=[]):
        # Optimised tag writer for <c> cell number elements in the inner loop.

        attr = ''

        for key, value in attributes:
            value = self._escape_attributes(value)
            attr = attr + ' %s="%s"' % (key, value)

        self.fh.write("""<c%s><v>%s</v></c>""" % (attr, str(number)))

    def _xml_formula_element(self, formula, result, attributes=[]):
        # Optimised tag writer for <c> cell formula elements in the inner loop.

        attr = ''

        for key, value in attributes:
            value = self._escape_attributes(value)
            attr = attr + ' %s="%s"' % (key, value)

        self.fh.write("""<c%s><f>%s</f><v>%s</v></c>""" 
                      % (attr, formula, str(result)))

    def _escape_attributes(self, str):
        # Escape XML characters in attributes.

        if not self.escapes.match(str):
            return str

        str = str.replace('&', '&amp;')
        str = str.replace('"', '&quot;')
        str = str.replace('<', '&lt;')
        str = str.replace('>', '&gt;')

        return str

    def _escape_data(self, str):
        # Escape XML characters in attributes.

        if not self.escapes.match(str):
            return str

        str = str.replace('&', '&amp;')
        str = str.replace('<', '&lt;')
        str = str.replace('>', '&gt;')

        return str
