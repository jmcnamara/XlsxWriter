###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from datetime import datetime
from ..helperfunctions import _xml_to_list
from ...core import Core


class TestAssembleCore(unittest.TestCase):
    """
    Test assembling a complete Core file.

    """
    def test_assemble_xml_file(self):
        """Test writing an Core file."""
        self.maxDiff = None

        fh = StringIO()
        core = Core()
        core._set_filehandle(fh)

        properties = {
            'title': 'This is an example spreadsheet',
            'subject': 'With document properties',
            'author': 'John McNamara',
            'manager': 'Dr. Heinz Doofenshmirtz',
            'company': 'of Wolves',
            'category': 'Example spreadsheets',
            'keywords': 'Sample, Example, Properties',
            'comments': 'Created with Python and XlsxWriter',
            'status': 'Quo',
            'created': datetime(2011, 4, 6, 19, 45, 15),
        }

        core._set_properties(properties)

        core._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                  <dc:title>This is an example spreadsheet</dc:title>
                  <dc:subject>With document properties</dc:subject>
                  <dc:creator>John McNamara</dc:creator>
                  <cp:keywords>Sample, Example, Properties</cp:keywords>
                  <dc:description>Created with Python and XlsxWriter</dc:description>
                  <cp:lastModifiedBy>John McNamara</cp:lastModifiedBy>
                  <dcterms:created xsi:type="dcterms:W3CDTF">2011-04-06T19:45:15Z</dcterms:created>
                  <dcterms:modified xsi:type="dcterms:W3CDTF">2011-04-06T19:45:15Z</dcterms:modified>
                  <cp:category>Example spreadsheets</cp:category>
                  <cp:contentStatus>Quo</cp:contentStatus>
                </cp:coreProperties>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
