###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2015, John McNamara, jmcnamara@cpan.org
#

import os
from zipfile import ZipFile
import unittest
from ... import workbook as wb
from ...workbook import Workbook


class TestExpandUser(unittest.TestCase):
    """
    Test storing workbook in a home directory.

    """

    def setUp(self):
        wb.ZipFile = ZipFileMock

    def tearDown(self):
        wb.ZipFile = ZipFile

    def test_expanduser(self):
        home_path = '~/tmp/xlsxwriter_test/book.xlsx'
        expanded_path = os.path.expanduser(home_path)
        self.assertNotEqual(home_path, expanded_path)
        workbook = Workbook(home_path)
        workbook.close()
        self.assertEqual(ZipFileMock.calls, [expanded_path])


class ZipFileMock(object):

    calls = []

    def __init__(self, file, *args, **kwargs):
        ZipFileMock.calls.append(file)

    def write(self, *args, **kwargs):
        pass

    def close(self):
        pass
