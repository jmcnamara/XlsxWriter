###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from ..compatibility import StringIO
from ...workbook import Workbook


class TestWriteDefinedNames(unittest.TestCase):
    """
    Test the Workbook _write_defined_names() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_defined_names(self):
        """Test the _write_defined_names() method"""

        self.workbook.defined_names = [['_xlnm.Print_Titles', 0, 'Sheet1!$1:$1', 0]]

        self.workbook._write_defined_names()

        exp = """<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName></definedNames>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    # TODO. Add define_name() test when ported.

    def tearDown(self):
        self.workbook.fileclosed = 1

if __name__ == '__main__':
    unittest.main()
