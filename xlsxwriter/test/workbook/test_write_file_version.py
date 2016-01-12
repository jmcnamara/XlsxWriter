###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...workbook import Workbook


class TestWriteFileVersion(unittest.TestCase):
    """
    Test the Workbook _write_file_version() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_file_version(self):
        """Test the _write_file_version() method"""

        self.workbook._write_file_version()

        exp = """<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
