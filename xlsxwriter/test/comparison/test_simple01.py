###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...workbook import Workbook


class TestCreateXLSXFile(unittest.TestCase):
    """
    Test TODO.

    """
    def test_create_file(self):
        """Test TODO."""
        self.maxDiff = None

        workbook = Workbook('simple01.xlsx')
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Hello')
        worksheet.write_number(1, 0, 123)

        workbook.close()

        exp = 1
        got = 1
        self.assertEqual(got, exp)


if __name__ == '__main__':
    unittest.main()
