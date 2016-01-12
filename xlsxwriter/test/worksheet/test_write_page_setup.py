###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWritePageSetup(unittest.TestCase):
    """
    Test the Worksheet _write_page_setup() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_page_setup_none(self):
        """Test the _write_page_setup() method. Without any page setup"""

        self.worksheet._write_page_setup()

        exp = ''
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_setup_landscape(self):
        """Test the _write_page_setup() method. With set_landscape()"""

        self.worksheet.set_landscape()

        self.worksheet._write_page_setup()

        exp = """<pageSetup orientation="landscape"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_setup_portrait(self):
        """Test the _write_page_setup() method. With set_portrait()"""

        self.worksheet.set_portrait()

        self.worksheet._write_page_setup()

        exp = """<pageSetup orientation="portrait"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_setup_paper(self):
        """Test the _write_page_setup() method. With set_paper()"""

        self.worksheet.set_paper(9)

        self.worksheet._write_page_setup()

        exp = """<pageSetup paperSize="9" orientation="portrait"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_setup_print_across(self):
        """Test the _write_page_setup() method. With print_across()"""

        self.worksheet.print_across()

        self.worksheet._write_page_setup()

        exp = """<pageSetup pageOrder="overThenDown" orientation="portrait"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
