###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import os
import unittest
from typing import Any, Dict, List

from .helperfunctions import _compare_xlsx_files


class ExcelComparisonTest(unittest.TestCase):
    """
    Test class for comparing a file created by XlsxWriter against a file
    created by Excel.

    """

    def __init__(self, *args: Any) -> None:
        """
        Initialize the ExcelComparisonTest instance.

        Args:
            *args: Variable arguments passed to unittest.TestCase.
        """
        super().__init__(*args)

        # pylint: disable=invalid-name
        self.maxDiff: None = None

        self.got_filename: str = ""
        self.exp_filename: str = ""
        self.ignore_files: List[str] = []
        self.ignore_elements: Dict[str, Any] = {}
        self.txt_filename: str = ""
        self.delete_output: bool = True

        # Set the paths for the test files.
        self.test_dir: str = "xlsxwriter/test/comparison/"
        self.vba_dir: str = self.test_dir + "xlsx_files/"
        self.image_dir: str = self.test_dir + "images/"
        self.theme_dir: str = self.test_dir + "themes/"
        self.output_dir: str = self.test_dir + "output/"

    def set_filename(self, filename: str) -> None:
        """
        Set the filenames for the Excel comparison test.

        Args:
            filename (str): The base filename for the test files.
        """
        # The reference Excel generated file.
        self.exp_filename = self.test_dir + "xlsx_files/" + filename

        # The generated XlsxWriter file.
        self.got_filename = self.output_dir + "py_" + filename

    def set_text_file(self, filename: str) -> None:
        """
        Set the filename and path for text files used in tests.

        Args:
            filename (str): The name of the text file.
        """
        # Set the filename and path for text files used in tests.
        self.txt_filename = self.test_dir + "xlsx_files/" + filename

    def assertExcelEqual(self) -> None:  # pylint: disable=invalid-name
        """
        Compare the generated file with the reference Excel file.

        Raises:
            AssertionError: If the files are not equivalent.
        """
        # Compare the generate file and the reference Excel file.
        got, exp = _compare_xlsx_files(
            self.got_filename,
            self.exp_filename,
            self.ignore_files,
            self.ignore_elements,
        )

        self.assertEqual(exp, got)

    def tearDown(self) -> None:
        """
        Clean up after each test by removing temporary files.
        Raises:
            OSError: If there is an error deleting the file.
        """
        # Cleanup by removing the temp excel file created for testing.
        if self.delete_output and os.path.exists(self.got_filename):
            os.remove(self.got_filename)
