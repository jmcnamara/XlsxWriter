###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
import os
from ...workbook import Workbook
from ..helperfunctions import _compare_xlsx_files


class TestCreateXLSXFile(unittest.TestCase):
    """
    Test TODO.

    """
    def test_create_file(self):
        """Test TODO."""
        self.maxDiff = None

        filename = 'simple02.xlsx'
        test_dir = 'xlsxwriter/test/comparison/'
        got_filename = test_dir + '_test_' + filename
        exp_filename = test_dir + 'xlsx_files/' + filename

        ignore_members = []
        ignore_elements = {}

        workbook = Workbook(got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet('Data Sheet')
        worksheet3 = workbook.add_worksheet()

        bold = workbook.add_format({'bold': 1})

        worksheet1.write_string(0, 0, 'Foo')
        worksheet1.write_number(1, 0, 123)

        worksheet3.write_string(1, 1, 'Foo')
        worksheet3.write_string(1, 2, 'Bar', bold)
        worksheet3.write_number(3, 3, 234)

        workbook.close()

        exp, got = _compare_xlsx_files(got_filename,
                                       exp_filename,
                                       ignore_members,
                                       ignore_elements)

        self.assertEqual(got, exp)

        # Cleanup.
        os.remove(got_filename)


if __name__ == '__main__':
    unittest.main()
