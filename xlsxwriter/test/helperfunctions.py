###############################################################################
#
# Helper functions for testing XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import re


class MockFormat(object):
    """
    Simple mocked Format object for testing.

    """
    def __init__(self, xf_index):
        self.xf_index = xf_index

    def _get_xf_index(self):
        return self.xf_index


def _xml_to_list(xml_str):
    # Convert test generated XML strings into lists for comparison testing.

    # Split the XML string at tag boundaries.
    parser = re.compile(r'>\s*<')
    elements = parser.split(xml_str.strip())

    # Add back the removed brackets.
    for index, element in enumerate(elements):
        if not element[0] == '<':
            elements[index] = '<' + elements[index]
        if not element[-1] == '>':
            elements[index] = elements[index] + '>'

    return elements
