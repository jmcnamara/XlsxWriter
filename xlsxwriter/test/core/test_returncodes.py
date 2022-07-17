###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
import datetime
import itertools
import warnings
import tempfile
import os
from ...returncodes import ReturnCode
from ...workbook import Workbook


class TestReturnCodes(unittest.TestCase):
    """
    Test return codes from the different modules.

    """

    _string_conversion_expected_result = "No error"
    _testing_image_path = 'xlsxwriter/test/comparison/images/logo.png'
    _testing_vba_bin_path = 'xlsxwriter/test/comparison/xlsx_files/vbaProject02.bin'

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

        self.worksheet = self.workbook.add_worksheet()

        self.max_col = self.worksheet.xls_colmax
        self.max_row = self.worksheet.xls_rowmax

        self.bold = self.workbook.add_format({'bold': True})

    ###########################################################################
    #
    # Helper functions
    #
    ###########################################################################

    def _test_no_error(self, func):
        """Test the no error return code of func

        Args:
            func: Function with prototype def func()
        """

        exp = ReturnCode.XW_NO_ERROR

        got = func()
        self.assertEqual(got, exp)

    def _test_cell_out_of_range(self, func, valid_r=0, valid_c=0):
        """Test the out of range return code of func

        This tests only a single cell value

        Args:
            func: Function with prototype def func(r,c), where r is the row
                index and c the column index
            valid_r (int, optional): Value for the valid row to be used in the
                                     test when necessary. Defaults to 0.
            other_c (int, optional): Value for the valid column to be used in
                                     the test when necessary. Defaults to 0.
        """

        exp = ReturnCode.XW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE

        got = func(valid_r, self.max_col)
        self.assertEqual(got, exp)

        got = func(self.max_row, valid_c)
        self.assertEqual(got, exp)

        got = func(self.max_row, self.max_col)
        self.assertEqual(got, exp)

    def _test_range_out_of_range(self, func, valid_r=0, valid_c=0):
        """Test the out of range return code of func

        This can be used for functions accepting a range of cells.

        Args:
            func: Function with prototype def func(r1,c1,r2,c2):, where (r1,c1)
                  is the starting cell and (r2,c2) is the ending cell of the
                  range to test
            valid_r (int, optional): Value for the valid row to be used in the
                                     test when necessary. Defaults to 0.
            other_c (int, optional): Value for the valid column to be used in
                                     the test when necessary. Defaults to 0.
        """

        exp = ReturnCode.XW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE

        # There are 16 cases; let's write them programmatically
        for shall_be_valid in itertools.product([True, False], repeat=4):
            # Skip all valid cases
            if all(shall_be_valid):
                continue

            r1 = valid_r if shall_be_valid[0] else self.max_row
            c1 = valid_c if shall_be_valid[1] else self.max_col
            r2 = valid_r if shall_be_valid[2] else self.max_row
            c2 = valid_c if shall_be_valid[3] else self.max_col

            got = func(r1, c1, r2, c2)
            self.assertEqual(got, exp)

    def _test_max_string_length(self, func):
        """Test the max string length return code of func

        Args:
            func: Function with prototype def func(s), where s is the string to
                  be written
        """

        long_string = " " * (self.worksheet.xls_strmax + 1)
        exp = ReturnCode.XW_ERROR_MAX_STRING_LENGTH_EXCEEDED
        got = func(long_string)

        self.assertEqual(got, exp)

    ###########################################################################
    #
    # Test string conversion
    #
    ###########################################################################

    def test_returncode_to_string_implicit(self):
        """Test converting a return code to string.

        Implicit conversion (using StrEnum underlying str representation)
        """
        retcode = ReturnCode.XW_NO_ERROR
        got = retcode

        self.assertEqual(got, self._string_conversion_expected_result)

    def test_returncode_to_string_format(self):
        """Test converting a return code to string.

        Using str.format() conversion
        """
        retcode = ReturnCode.XW_NO_ERROR
        got = "{0}".format(retcode)

        self.assertEqual(got, self._string_conversion_expected_result)

    def test_returncode_to_string_fstring(self):
        """Test converting a return code to string.

        Using f-string conversion
        """
        retcode = ReturnCode.XW_NO_ERROR
        got = f'{retcode}'

        self.assertEqual(got, self._string_conversion_expected_result)

    ###########################################################################
    #
    # Test workbook object return codes
    #
    ###########################################################################

    def test_workbook_add_vba_project_no_error(self):
        """worksheet.add_vba_project() returns XW_NO_ERROR
        """

        def func():
            return self.workbook.add_vba_project(self._testing_vba_bin_path)
        self._test_no_error(func)

    def test_workbook_add_vba_project_vba_file_not_found(self):
        """worksheet.add_vba_project() returns VBA_FILE_NOT_FOUND
        """

        exp = ReturnCode.XW_ERROR_VBA_FILE_NOT_FOUND

        # Ignore the warning "VBA project binary file not found"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            got = self.workbook.add_vba_project('unknown.bin')
            self.assertEqual(got, exp)

    def test_workbook_set_custom_property_no_error(self):
        """worksheet.set_custom_property() returns XW_NO_ERROR
        """

        def func():
            return self.workbook.set_custom_property('test', True)
        self._test_no_error(func)

    def test_worksheet_set_custom_property_incorrect_parameter_or_option(self):
        """worksheet.set_custom_property() returns INCORRECT_PARAMETER_OR_OPTION
        """

        # Ignore the warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INCORRECT_PARAMETER_OR_OPTION

            # 01) The name and value parameters must be non-None
            got = self.workbook.set_custom_property(None, None)
            self.assertEqual(got, exp)

    def test_workbook_define_name_no_error(self):
        """worksheet.define_name() returns XW_NO_ERROR
        """

        def func():
            return self.workbook.define_name('Test', 'A1')
        self._test_no_error(func)

    def test_worksheet_define_name_parameter_or_option(self):
        """worksheet.define_name() returns INCORRECT_PARAMETER_OR_OPTION
        """

        # Ignore the warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INCORRECT_PARAMETER_OR_OPTION

            # 01) Unknown sheet name
            got = self.workbook.define_name('NoSheet!Test', 'A1')
            self.assertEqual(got, exp)

            # 02) Invalid Excel characters
            got = self.workbook.define_name('.', 'A1')
            self.assertEqual(got, exp)

            # 03) Name looks like a cell name
            got = self.workbook.define_name('A0', 'A1')
            self.assertEqual(got, exp)

            # 04) Invalid name like a RC cell ref
            got = self.workbook.define_name('R1C1', 'A1')
            self.assertEqual(got, exp)

    ###########################################################################
    #
    # Test worksheet object return codes
    #
    ###########################################################################

    def test_worksheet_write_string_no_error(self):
        """worksheet.write_string() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_string(0, 0, '')
        self._test_no_error(func)

    def test_worksheet_write_string_out_of_range(self):
        """worksheet.write_string() returns INDEX_OUT_OF_RANGE

        This is already tested in worksheet.test_range_return_values, but for
        completeness sake it is also repeated here
        """

        def func(r, c):
            return self.worksheet.write_string(r, c, '')
        self._test_cell_out_of_range(func)

    def test_worksheet_write_string_max_string_length(self):
        """worksheet.write_string() returns MAX_STRING_LENGTH_EXCEEDED
        """

        def func(s):
            return self.worksheet.write_string(0, 0, s)
        self._test_max_string_length(func)

    def test_worksheet_write_number_no_error(self):
        """worksheet.write_number() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_number(0, 0, 0)
        self._test_no_error(func)

    def test_worksheet_write_number_out_of_range(self):
        """worksheet.write_number() returns INDEX_OUT_OF_RANGE

        This is already tested in worksheet.test_range_return_values, but for
        completeness sake it is also repeated here
        """

        def func(r, c):
            return self.worksheet.write_number(r, c, 3)
        self._test_cell_out_of_range(func)

    def test_worksheet_write_blank_no_error(self):
        """worksheet.write_blank() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_blank(0, 0, None, 'format')
        self._test_no_error(func)

        # Test that without format even out of range is not returned
        def func():
            return self.worksheet.write_blank(0, self.max_col, None)
        self._test_no_error(func)

    def test_worksheet_write_blank_out_of_range(self):
        """worksheet.write_blank() returns INDEX_OUT_OF_RANGE

        This is already tested in worksheet.test_range_return_values, but for
        completeness sake it is also repeated here
        """

        def func(r, c):
            return self.worksheet.write_blank(r, c, None, 'format')
        self._test_cell_out_of_range(func)

    def test_worksheet_write_formula_no_error(self):
        """worksheet.write_formula() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_formula(0, 0, '=B2')
        self._test_no_error(func)

    def test_worksheet_write_formula_out_of_range(self):
        """worksheet.write_formula() returns INDEX_OUT_OF_RANGE

        This is already tested in worksheet.test_range_return_values, but for
        completeness sake it is also repeated here
        """

        def func(r, c):
            return self.worksheet.write_formula(r, c, '=B2')
        self._test_cell_out_of_range(func)

    def test_worksheet_write_formula_none_or_empty(self):
        """worksheet.write_formula() returns XW_ERROR_FORMULA_CANT_BE_NONE_OR_EMPTY
        """

        # Ignore the warning "Formula can't be None or empty"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_FORMULA_CANT_BE_NONE_OR_EMPTY

            got = self.worksheet.write_formula(0, 0, None)
            self.assertEqual(got, exp)

            got = self.worksheet.write_formula(0, 0, '')
            self.assertEqual(got, exp)

    def test_worksheet_write_array_formula_no_error(self):
        """worksheet.write_array_formula() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_array_formula(0, 0, 0, 0, '')
        self._test_no_error(func)

    def test_worksheet_write_array_formula_out_of_range(self):
        """worksheet.write_array_formula() returns INDEX_OUT_OF_RANGE

        This is already tested in worksheet.test_range_return_values, but for
        completeness sake it is also repeated here
        """

        def func(r1, c1, r2, c2):
            return self.worksheet.write_array_formula(r1, c1, r2, c2, '')
        self._test_range_out_of_range(func)

    def test_worksheet_write_dynamic_array_formula_no_error(self):
        """worksheet.write_dynamic_array_formula() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_dynamic_array_formula(0, 0, 0, 0, '')
        self._test_no_error(func)

    def test_worksheet_write_dynamic_array_formula_out_of_range(self):
        """worksheet.write_dynamic_array_formula() returns INDEX_OUT_OF_RANGE
        """

        def func(r1, c1, r2, c2):
            return self.worksheet.write_dynamic_array_formula(r1, c1, r2, c2, '')
        self._test_range_out_of_range(func)

    def test_worksheet_write_datetime_no_error(self):
        """worksheet.write_datetime() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_datetime(0, 0, datetime.datetime.now())
        self._test_no_error(func)

    def test_worksheet_write_datetime_out_of_range(self):
        """worksheet.write_datetime() returns INDEX_OUT_OF_RANGE
        """

        def func(r, c):
            return self.worksheet.write_datetime(r, c, datetime.datetime.now())
        self._test_cell_out_of_range(func)

    def test_worksheet_write_boolean_no_error(self):
        """worksheet.write_boolean() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_boolean(0, 0, True)
        self._test_no_error(func)

    def test_worksheet_write_boolean_out_of_range(self):
        """worksheet.write_boolean() returns INDEX_OUT_OF_RANGE
        """

        def func(r, c):
            return self.worksheet.write_boolean(r, c, True)
        self._test_cell_out_of_range(func)

    def test_worksheet_write_url_no_error(self):
        """worksheet.write_url() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_url(0, 0, '')
        self._test_no_error(func)

    def test_worksheet_write_url_out_of_range(self):
        """worksheet.write_url() returns INDEX_OUT_OF_RANGE
        """

        def func(r, c):
            return self.worksheet.write_url(r, c, '')
        self._test_cell_out_of_range(func)

    def test_worksheet_write_url_max_string_length(self):
        """worksheet.write_url() returns MAX_STRING_LENGTH_EXCEEDED
        """

        # Ignore the warning "Ignoring URL since it exceeds Excel's string limit"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            def func(s):
                return self.worksheet.write_url(0, 0, s)
            self._test_max_string_length(func)

    def test_worksheet_write_url_max_url_length(self):
        """worksheet.write_url() returns MAX_URL_LENGTH_EXCEEDED
        """

        # Ignore the warning "Ignoring URL '%s' with link or location/anchor"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            # Set a forced value for URL length
            forced_url_length = 500

            self.worksheet.max_url_length = forced_url_length

            long_url = " " * (forced_url_length + 1)

            exp = ReturnCode.XW_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED
            got = self.worksheet.write_url(0, 0, long_url)

            self.assertEqual(got, exp)

    def test_worksheet_write_url_max_number_urls(self):
        """worksheet.write_url() returns MAX_NUMBER_URLS_EXCEEDED
        """

        # Ignore the warning "Ignoring URL '%s' with link or location/anchor"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            # Set a forced value for URL length
            max_urls_per_worksheet = 65530

            exp = ReturnCode.XW_NO_ERROR
            for _ in range(max_urls_per_worksheet):
                got = self.worksheet.write_url(0, 0, '')
                self.assertEqual(got, exp)

            exp = ReturnCode.XW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED
            got = self.worksheet.write_url(0, 0, '')

            self.assertEqual(got, exp)

    def test_worksheet_write_rich_string_no_error(self):
        """worksheet.write_rich_string() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_rich_string(0, 0, 'a', self.bold, 'b')
        self._test_no_error(func)

    def test_worksheet_write_rich_string_out_of_range(self):
        """worksheet.write_rich_string() returns INDEX_OUT_OF_RANGE
        """

        def func(r, c):
            return self.worksheet.write_rich_string(r, c, 'a', self.bold, 'b')
        self._test_cell_out_of_range(func)

    def test_worksheet_write_rich_string_max_string_length(self):
        """worksheet.write_rich_string() returns MAX_STRING_LENGTH_EXCEEDED
        """

        # Ignore the warning "Ignoring URL since it exceeds Excel's string limit"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            # Removing the last char from s since the function will concatenate 'a'
            def func(s):
                return self.worksheet.write_rich_string(0, 0, s[:-1], self.bold, 'a')
            self._test_max_string_length(func)

    def test_worksheet_write_rich_string_2_consecutive_formats(self):
        """worksheet.write_rich_string() returns 2_CONSECUTIVE_FORMATS
        """

        # Ignore the warning "Excel doesn't allow 2 consecutive formats"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_2_CONSECUTIVE_FORMATS

            got = self.worksheet.write_rich_string(0, 0, 'a', self.bold, self.bold, 'b')
            self.assertEqual(got, exp)

    def test_worksheet_write_rich_string_empty_string_used(self):
        """worksheet.write_rich_string() returns EMPTY_STRING_USED
        """

        # Ignore the warning "Excel doesn't allow empty strings in rich strings"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_EMPTY_STRING_USED

            got = self.worksheet.write_rich_string(0, 0, '', self.bold, 'b')
            self.assertEqual(got, exp)

    def test_worksheet_write_rich_string_insufficient_parameters(self):
        """worksheet.write_rich_string() returns INSUFFICIENT_PARAMETERS
        """

        # Ignore the warning "You must specify more than 2 format/fragments"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INSUFFICIENT_PARAMETERS

            got = self.worksheet.write_rich_string(0, 0, 'a')
            self.assertEqual(got, exp)

    def test_worksheet_write_row_no_error(self):
        """worksheet.write_row() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_row(0, 0, [0, 1, 2])
        self._test_no_error(func)

    def test_worksheet_write_column_no_error(self):
        """worksheet.write_column() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_column(0, 0, [0, 1, 2])
        self._test_no_error(func)

    def test_worksheet_insert_image_no_error(self):
        """worksheet.insert_image() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.insert_image(0, 0, self._testing_image_path)
        self._test_no_error(func)

    def test_worksheet_insert_image_out_of_range(self):
        """worksheet.insert_image() returns INDEX_OUT_OF_RANGE
        """

        # Ignore the warning "Cannot insert image"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            def func(r, c):
                return self.worksheet.insert_image(r, c, self._testing_image_path)
            self._test_cell_out_of_range(func)

    def test_worksheet_insert_image_file_not_found(self):
        """worksheet.insert_image() returns XW_ERROR_IMAGE_FILE_NOT_FOUND
        """

        # Ignore the warning "Image file not found"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_IMAGE_FILE_NOT_FOUND

            got = self.worksheet.insert_image(0, 0, 'xlsxwriter/nonexisting.png')
            self.assertEqual(got, exp)

    def test_worksheet_insert_textbox_no_error(self):
        """worksheet.insert_textbox() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.insert_textbox(0, 0, '')
        self._test_no_error(func)

    def test_worksheet_insert_textbox_out_of_range(self):
        """worksheet.insert_textbox() returns INDEX_OUT_OF_RANGE
        """

        # Ignore the warning "Cannot insert textbox"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            def func(r, c):
                return self.worksheet.insert_textbox(r, c, '')
            self._test_cell_out_of_range(func)

    def test_worksheet_insert_chart_no_error(self):
        """worksheet.insert_chart() returns XW_NO_ERROR
        """

        chart = self.workbook.add_chart({'type': 'column'})

        def func():
            return self.worksheet.insert_chart(0, 0, chart)
        self._test_no_error(func)

    def test_worksheet_insert_chart_out_of_range(self):
        """worksheet.insert_chart() returns INDEX_OUT_OF_RANGE
        """

        # Ignore the warning "Cannot insert chart"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            chart = self.workbook.add_chart({'type': 'column'})

            def func(r, c):
                return self.worksheet.insert_chart(r, c, chart)
            self._test_cell_out_of_range(func)

    def test_worksheet_insert_chart_none(self):
        """worksheet.insert_chart() returns None
        """

        chart = self.workbook.add_chart({'type': 'column'})

        # Insert it the first time; no error
        def func():
            return self.worksheet.insert_chart(0, 0, chart)
        self._test_no_error(func)

        # Try to insert the same chart twice, but ignore the warning
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            got = self.worksheet.insert_chart(0, 0, chart)
            self.assertIsNone(got)

    def test_worksheet_write_comment_no_error(self):
        """worksheet.write_comment() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.write_comment(0, 0, '')
        self._test_no_error(func)

    def test_worksheet_write_comment_out_of_range(self):
        """worksheet.write_comment() returns INDEX_OUT_OF_RANGE
        """

        def func(r, c):
            return self.worksheet.write_comment(r, c, '')
        self._test_cell_out_of_range(func)

    def test_worksheet_write_comment_max_string_length(self):
        """worksheet.write_comment() returns MAX_STRING_LENGTH_EXCEEDED
        """

        def func(s):
            return self.worksheet.write_comment(0, 0, s)
        self._test_max_string_length(func)

    def test_worksheet_set_background_no_error(self):
        """worksheet.set_background() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.set_background(self._testing_image_path)
        self._test_no_error(func)

    def test_worksheet_set_background_file_not_found(self):
        """worksheet.set_background() returns XW_ERROR_IMAGE_FILE_NOT_FOUND
        """

        # Ignore the warning "Image file not found"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_IMAGE_FILE_NOT_FOUND

            got = self.worksheet.set_background('xlsxwriter/nonexisting.png')
            self.assertEqual(got, exp)

    def test_worksheet_set_column_no_error(self):
        """worksheet.set_column() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.set_column(0, 0, 10)
        self._test_no_error(func)

    def test_worksheet_set_column_out_of_range(self):
        """worksheet.set_column() returns INDEX_OUT_OF_RANGE
        """

        def func(r, c):
            return self.worksheet.set_column(r, c, 10)
        self._test_cell_out_of_range(func)

    def test_worksheet_set_column_pixels_no_error(self):
        """worksheet.set_column_pixels() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.set_column_pixels(0, 0, 10)
        self._test_no_error(func)

    def test_worksheet_set_column_pixels_out_of_range(self):
        """worksheet.set_column_pixels() returns INDEX_OUT_OF_RANGE
        """

        def func(r, c):
            return self.worksheet.set_column_pixels(r, c, 10)
        self._test_cell_out_of_range(func)

    def test_worksheet_set_row_no_error(self):
        """worksheet.set_row() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.set_row(0, 10)
        self._test_no_error(func)

    def test_worksheet_set_row_out_of_range(self):
        """worksheet.set_row() returns INDEX_OUT_OF_RANGE
        """

        exp = ReturnCode.XW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE

        got = self.worksheet.set_row(self.max_row, 10)
        self.assertEqual(got, exp)

    def test_worksheet_set_row_pixels_no_error(self):
        """worksheet.set_row_pixels() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.set_row_pixels(0, 10)
        self._test_no_error(func)

    def test_worksheet_set_row_pixels_out_of_range(self):
        """worksheet.set_row_pixels() returns INDEX_OUT_OF_RANGE
        """

        exp = ReturnCode.XW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE

        got = self.worksheet.set_row_pixels(self.max_row, 10)
        self.assertEqual(got, exp)

    def test_worksheet_merge_range_no_error(self):
        """worksheet.merge_range() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.merge_range(0, 0, 0, 1, '')
        self._test_no_error(func)

    def test_worksheet_merge_range_out_of_range(self):
        """worksheet.merge_range() returns INDEX_OUT_OF_RANGE

        This is already tested in worksheet.test_range_return_values, but for
        completeness sake it is also repeated here
        """

        def func(r1, c1, r2, c2):
            return self.worksheet.merge_range(r1 - 1 if r1 == r2 else r1, c1 - 1 if c1 == c2 else c1, r2, c2, '')
        self._test_range_out_of_range(func)

    def test_worksheet_merge_range_none(self):
        """worksheet.merge_range() returns None
        """

        # Ignore the warning "Can't merge single cell"
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            got = self.worksheet.merge_range(0, 0, 0, 0, '')
            self.assertIsNone(got)

    def test_worksheet_data_validation_no_error(self):
        """worksheet.data_validation() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'list', 'source': ['a']})
        self._test_no_error(func)

    def test_worksheet_data_validation_out_of_range(self):
        """worksheet.data_validation() returns INDEX_OUT_OF_RANGE
        """

        def func(r1, c1, r2, c2):
            return self.worksheet.data_validation(r1, c1, r2, c2, {'validate': 'list', 'source': ['a']})
        self._test_range_out_of_range(func)

    def test_worksheet_data_validation_incorrect_parameter_or_option(self):
        """worksheet.data_validation() returns INCORRECT_PARAMETER_OR_OPTION
        """

        # Ignore the warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INCORRECT_PARAMETER_OR_OPTION

            # 01) Unknown parameter
            got = self.worksheet.data_validation(0, 0, 0, 0, {'unknown_param': 0})
            self.assertEqual(got, exp)

            # 02) Parameter 'validate' is required
            got = self.worksheet.data_validation(0, 0, 0, 0, {})
            self.assertEqual(got, exp)

            # 03) Unknown validation type
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'unknown'})
            self.assertEqual(got, exp)

            # 04) No action is required
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'any'})
            self.assertEqual(got, exp)

            # 05) 'criteria' is a required parameter (for some validate)
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'integer'})
            self.assertEqual(got, exp)

            # 06) Unknown criteria type
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'integer', 'criteria': 'unknown'})
            self.assertEqual(got, exp)

            # 07) 'Between' and 'Not between' criteria require 2 values
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'integer', 'criteria': 'between'})
            self.assertEqual(got, exp)

            # 08) Unknown criteria type for parameter 'error_type'
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'any', 'input_title': ' ', 'error_type': 'unknown'})
            self.assertEqual(got, exp)

            # 09) Length of input title
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'any', 'input_title': ' ' * 33})
            self.assertEqual(got, exp)

            # 10) Length of error title
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'any', 'input_title': ' ', 'error_title': ' ' * 33})
            self.assertEqual(got, exp)

            # 11) Length of input message
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'any', 'input_title': ' ', 'input_message': ' ' * 256})
            self.assertEqual(got, exp)

            # 12) Length of error message
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'any', 'input_title': ' ', 'error_message': ' ' * 256})
            self.assertEqual(got, exp)

            # 13) Length of list items
            got = self.worksheet.data_validation(0, 0, 0, 0, {'validate': 'list', 'source': [' ' * 256]})
            self.assertEqual(got, exp)

    def test_worksheet_conditional_format_no_error(self):
        """worksheet.conditional_format() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.conditional_format(0, 0, 0, 0, {'type': 'cell', 'criteria': '>=', 'value': 50, 'format': self.bold})
        self._test_no_error(func)

    def test_worksheet_conditional_format_out_of_range(self):
        """worksheet.conditional_format() returns INDEX_OUT_OF_RANGE
        """

        def func(r1, c1, r2, c2):
            return self.worksheet.conditional_format(r1, c1, r2, c2, {'type': 'cell', 'criteria': '>=', 'value': 50, 'format': self.bold})
        self._test_range_out_of_range(func)

    def test_worksheet_conditional_format_incorrect_parameter_or_option(self):
        """worksheet.conditional_format() returns INCORRECT_PARAMETER_OR_OPTION
        """

        # Ignore the warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INCORRECT_PARAMETER_OR_OPTION

            # 01) Unknown parameter
            got = self.worksheet.conditional_format(0, 0, 0, 0, {'unknown_param': 0})
            self.assertEqual(got, exp)

            # 02) Parameter 'type' is required
            got = self.worksheet.conditional_format(0, 0, 0, 0, {})
            self.assertEqual(got, exp)

            # 03) Unknown value for parameter 'type'
            got = self.worksheet.conditional_format(0, 0, 0, 0, {'type': 'unknown'})
            self.assertEqual(got, exp)

            # 04) Conditional format 'value' must be a datetime object.
            got = self.worksheet.conditional_format(0, 0, 0, 0, {'type': 'date', 'value': ''})
            self.assertEqual(got, exp)

            # 05) Conditional format 'minimum' must be a datetime object.
            got = self.worksheet.conditional_format(0, 0, 0, 0, {'type': 'date', 'minimum': ''})
            self.assertEqual(got, exp)

            # 06) Conditional format 'maximum' must be a datetime object.
            got = self.worksheet.conditional_format(0, 0, 0, 0, {'type': 'date', 'maximum': ''})
            self.assertEqual(got, exp)

            # 07) The 'icon_style' parameter must be specified
            got = self.worksheet.conditional_format(0, 0, 0, 0, {'type': 'icon_set'})
            self.assertEqual(got, exp)

            # 08) Unknown icon_style
            got = self.worksheet.conditional_format(0, 0, 0, 0, {'type': 'icon_set', 'icon_style': 'unknown'})
            self.assertEqual(got, exp)

    def test_worksheet_add_table_no_error(self):
        """worksheet.add_table() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.add_table(0, 0, 1, 0)
        self._test_no_error(func)

    def test_worksheet_add_table_not_supported_constant_memory(self):
        """worksheet.add_table() returns NOT_SUPPORTED_COSTANT_MEMORY
        """
        # Here we must use an actual file because the worksheet file handle is
        # not cleaned up correctly when using the StringIO

        (fd, filepath) = tempfile.mkstemp()
        os.close(fd)

        with Workbook(filepath, dict(constant_memory=True)) as tempworkbook:
            tempworksheet = tempworkbook.add_worksheet()
            with warnings.catch_warnings():
                warnings.simplefilter('ignore', category=UserWarning)

                exp = ReturnCode.XW_ERROR_NOT_SUPPORTED_COSTANT_MEMORY

                got = tempworksheet.add_table(0, 0, 1, 0)
                self.assertEqual(got, exp)

        os.unlink(filepath)

    def test_worksheet_add_table_out_of_range(self):
        """worksheet.add_table() returns INDEX_OUT_OF_RANGE
        """

        def func(r1, c1, r2, c2):
            return self.worksheet.add_table(r1, c1, r2, c2)
        self._test_range_out_of_range(func)

    def test_worksheet_add_table_incorrect_parameter_or_option(self):
        """worksheet.add_table() returns INCORRECT_PARAMETER_OR_OPTION
        """

        # Ignore the warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INCORRECT_PARAMETER_OR_OPTION

            # 01) Unknown parameter
            got = self.worksheet.add_table(0, 0, 0, 0, {'unknown_param': 0})
            self.assertEqual(got, exp)

            # 02) At least one data row
            got = self.worksheet.add_table(0, 0, 0, 0, {})
            self.assertEqual(got, exp)

            # 03) Name cannot contain spaces
            got = self.worksheet.add_table(0, 0, 1, 0, {'name': 'My Name'})
            self.assertEqual(got, exp)

            # 04) Invalid Excel characters
            got = self.worksheet.add_table(0, 0, 1, 0, {'name': '.'})
            self.assertEqual(got, exp)

            # 05) Name looks like a cell name
            got = self.worksheet.add_table(0, 0, 1, 0, {'name': 'A0'})
            self.assertEqual(got, exp)

            # 06) Invalid name like a RC cell ref
            got = self.worksheet.add_table(0, 0, 1, 0, {'name': 'R1C1'})
            self.assertEqual(got, exp)

            # 07) Duplicate header name
            got = self.worksheet.add_table(0, 0, 1, 1, {'columns': [{'header': 'a'}, {'header': 'a'}]})
            self.assertEqual(got, exp)

    def test_worksheet_add_sparkline_no_error(self):
        """worksheet.add_sparkline() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.add_sparkline(0, 0, {'range': 'A1:E1'})
        self._test_no_error(func)

    def test_worksheet_add_sparkline_out_of_range(self):
        """worksheet.add_sparkline() returns INDEX_OUT_OF_RANGE
        """

        def func(r, c):
            return self.worksheet.add_sparkline(r, c, {'range': 'A1:E1'})
        self._test_cell_out_of_range(func)

    def test_worksheet_add_sparkline_incorrect_parameter_or_option(self):
        """worksheet.add_sparkline() returns INCORRECT_PARAMETER_OR_OPTION
        """

        # Ignore the warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INCORRECT_PARAMETER_OR_OPTION

            # 01) Unknown parameter
            got = self.worksheet.add_sparkline(0, 0, {'unknown': 0})
            self.assertEqual(got, exp)

            # 02) Parameter 'range' is required
            got = self.worksheet.add_sparkline(0, 0, {})
            self.assertEqual(got, exp)

            # 03) Parameter 'type' must be 'line', 'column' or 'win_loss'
            got = self.worksheet.add_sparkline(0, 0, {'range': 'A1:E1', 'type': 'unknown'})
            self.assertEqual(got, exp)

            # 04) Must have the same number of location and range
            got = self.worksheet.add_sparkline(0, 0, {'range': ['A1:E1', 'E2:E3']})
            self.assertEqual(got, exp)

    def test_worksheet_unprotect_range_no_error(self):
        """worksheet.unprotect_range() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.unprotect_range('A1')
        self._test_no_error(func)

    def test_worksheet_unprotect_range_incorrect_parameter_or_option(self):
        """worksheet.unprotect_range() returns INCORRECT_PARAMETER_OR_OPTION
        """

        # Ignore Cell range must be specified
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INCORRECT_PARAMETER_OR_OPTION

            got = self.worksheet.unprotect_range(None)
            self.assertEqual(got, exp)

    def test_worksheet_insert_button_no_error(self):
        """worksheet.insert_button() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.insert_button(0, 0)
        self._test_no_error(func)

    def test_worksheet_insert_button_out_of_range(self):
        """worksheet.insert_button() returns INDEX_OUT_OF_RANGE
        """

        # Ignore Cannot insert button
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            def func(r, c):
                return self.worksheet.insert_button(r, c)
            self._test_cell_out_of_range(func)

    def test_worksheet_print_area_no_error(self):
        """worksheet.print_area() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.print_area(0, 0, 10, 10)
        self._test_no_error(func)

    def test_worksheet_ignore_errors_no_error(self):
        """worksheet.ignore_errors() returns XW_NO_ERROR
        """

        def func():
            return self.worksheet.ignore_errors({'eval_error': True})
        self._test_no_error(func)

    def test_worksheet_ignore_errors_incorrect_parameter_or_option(self):
        """worksheet.ignore_errors() returns INCORRECT_PARAMETER_OR_OPTION
        """

        # Ignore the warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', category=UserWarning)

            exp = ReturnCode.XW_ERROR_INCORRECT_PARAMETER_OR_OPTION

            # 01) Parameter is None
            got = self.worksheet.ignore_errors()
            self.assertEqual(got, exp)

            # 02) Unknown parameter
            got = self.worksheet.ignore_errors({'unknown': 0})
            self.assertEqual(got, exp)
