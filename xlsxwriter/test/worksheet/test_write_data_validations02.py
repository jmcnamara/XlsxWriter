###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from datetime import date, time
from io import StringIO
from ...worksheet import Worksheet
from ..helperfunctions import _xml_to_list


class TestWriteDataValidations(unittest.TestCase):
    """
    Test the Worksheet _write_data_validations() method.

    """

    def setUp(self):
        self.maxDiff = None
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_data_validations_1(self):
        """
        Test 1 Integer between 1 and 10.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_2(self):
        """
        Test 2 Integer not between 1 and 10.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "not between",
                "minimum": 1,
                "maximum": 10,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_3(self):
        """
        Test 3,4,5 Integer == 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "equal to",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_4(self):
        """
        Test 3,4,5 Integer == 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "=",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_5(self):
        """
        Test 3,4,5 Integer == 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "==",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_6(self):
        """
        Test 6,7,8 Integer != 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "not equal to",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_7(self):
        """
        Test 6,7,8 Integer != 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "<>",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_8(self):
        """
        Test 6,7,8 Integer != 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "!=",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_9(self):
        """
        Test 9,10 Integer > 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "greater than",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_10(self):
        """
        Test 9,10 Integer > 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": ">",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_11(self):
        """
        Test 11,12 Integer < 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "less than",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_12(self):
        """
        Test 11,12 Integer < 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "<",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_13(self):
        """
        Test 13,14 Integer >= 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "greater than or equal to",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_14(self):
        """
        Test 13,14 Integer >= 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": ">=",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_15(self):
        """
        Test 15,16 Integer <= 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "less than or equal to",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_16(self):
        """
        Test 15,16 Integer <= 1.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "<=",
                "value": 1,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_17(self):
        """
        Test 17 Integer between 1 and 10 (same as test 1) + Ignore blank off.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "ignore_blank": 0,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_18(self):
        """
        Test 18 Integer between 1 and 10 (same as test 1) + Error style == warning.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "error_type": "warning",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_19(self):
        """
        Test 19 Integer between 1 and 10 (same as test 1) + Error style == info.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "error_type": "information",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" errorStyle="information" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_20(self):
        """
        Test 20 Integer between 1 and 10 (same as test 1)
                + input title.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "input_title": "Input title January",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Input title January" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_21(self):
        """
        Test 21 Integer between 1 and 10 (same as test 1)
                + input title.
                + input message.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "input_title": "Input title January",
                "input_message": "Input message February",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_22(self):
        """
        Test 22 Integer between 1 and 10 (same as test 1)
                + input title.
                + input message.
                + error title.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "input_title": "Input title January",
                "input_message": "Input message February",
                "error_title": "Error title March",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Error title March" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_23(self):
        """
        Test 23 Integer between 1 and 10 (same as test 1)
                + input title.
                + input message.
                + error title.
                + error message.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "input_title": "Input title January",
                "input_message": "Input message February",
                "error_title": "Error title March",
                "error_message": "Error message April",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_24(self):
        """
        Test 24 Integer between 1 and 10 (same as test 1)
                + input title.
                + input message.
                + error title.
                + error message.
                - input message box.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "input_title": "Input title January",
                "input_message": "Input message February",
                "error_title": "Error title March",
                "error_message": "Error message April",
                "show_input": 0,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showErrorMessage="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_25(self):
        """
        Test 25 Integer between 1 and 10 (same as test 1)
                + input title.
                + input message.
                + error title.
                + error message.
                - input message box.
                - error message box.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "input_title": "Input title January",
                "input_message": "Input message February",
                "error_title": "Error title March",
                "error_message": "Error message April",
                "show_input": 0,
                "show_error": 0,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_26(self):
        """
        Test 26 'Any' shouldn't produce a DV record if there are no messages.
        """
        self.worksheet.data_validation("B5", {"validate": "any"})

        self.worksheet._write_data_validations()

        exp = ""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_data_validations_27(self):
        """
        Test 27 Decimal = 1.2345
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "decimal",
                "criteria": "==",
                "value": 1.2345,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="decimal" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1.2345</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_28(self):
        """
        Test 28 List = a,bb,ccc
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "list",
                "source": ["a", "bb", "ccc"],
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>"a,bb,ccc"</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_29(self):
        """
        Test 29 List = a,bb,ccc, No dropdown
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "list",
                "source": ["a", "bb", "ccc"],
                "dropdown": 0,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>"a,bb,ccc"</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_30(self):
        """
        Test 30 List = $D$1:$D$5
        """
        self.worksheet.data_validation(
            "A1:A1",
            {
                "validate": "list",
                "source": "=$D$1:$D$5",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>$D$1:$D$5</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_31(self):
        """
        Test 31 Date = 39653 (2008-07-24)
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "date",
                "criteria": "==",
                "value": date(2008, 7, 24),
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="date" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39653</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_32(self):
        """
        Test 32 Date = 2008-07-25T
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "date",
                "criteria": "==",
                "value": date(2008, 7, 25),
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="date" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39654</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_33(self):
        """
        Test 33 Date between ranges.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "date",
                "criteria": "between",
                "minimum": date(2008, 1, 1),
                "maximum": date(2008, 12, 12),
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39448</formula1><formula2>39794</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_34(self):
        """
        Test 34 Time = 0.5 (12:00:00)
        """
        self.worksheet.data_validation(
            "B5:B5",
            {
                "validate": "time",
                "criteria": "==",
                "value": time(12),
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="time" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>0.5</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_35(self):
        """
        Test 35 Time = T12:00:00
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "time",
                "criteria": "==",
                "value": time(12, 0, 0),
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="time" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>0.5</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_36(self):
        """
        Test 36 Custom == 10.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "custom",
                "criteria": "==",
                "value": 10,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>10</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_37(self):
        """
        Test 37 Check the row/col processing: single A1 style cell.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_38(self):
        """
        Test 38 Check the row/col processing: single A1 style range.
        """
        self.worksheet.data_validation(
            "B5:B10",
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_39(self):
        """
        Test 39 Check the row/col processing: single (row, col) style cell.
        """
        self.worksheet.data_validation(
            4,
            1,
            4,
            1,
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_40(self):
        """
        Test 40 Check the row/col processing: single (row, col) style range.
        """
        self.worksheet.data_validation(
            4,
            1,
            9,
            1,
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_41(self):
        """
        Test 41 Check the row/col processing: multiple (row, col) style cells.
        """
        self.worksheet.data_validation(
            4,
            1,
            4,
            1,
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "other_cells": [[4, 3, 4, 3]],
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5 D5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_42(self):
        """
        Test 42 Check the row/col processing: multiple (row, col) style cells.
        """
        self.worksheet.data_validation(
            4,
            1,
            4,
            1,
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "other_cells": [[6, 1, 6, 1], [8, 1, 8, 1]],
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5 B7 B9"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_43(self):
        """
        Test 43 Check the row/col processing: multiple (row, col) style cells.
        """
        self.worksheet.data_validation(
            4,
            1,
            8,
            1,
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "other_cells": [[3, 3, 3, 3]],
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B9 D4"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_44(self):
        """
        Test 44 Multiple validations.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "integer",
                "criteria": ">",
                "value": 10,
            },
        )

        self.worksheet.data_validation(
            "C10",
            {
                "validate": "integer",
                "criteria": "<",
                "value": 10,
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="2"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>10</formula1></dataValidation><dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C10"><formula1>10</formula1></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_45(self):
        """
        Test 45 Test 'any' with input messages.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "any",
                "input_title": "Input title January",
                "input_message": "Input message February",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Input title January" prompt="Input message February" sqref="B5"/></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_46(self):
        """
        Test 46 Date between ranges with formulas.
        """
        self.worksheet.data_validation(
            "B5",
            {
                "validate": "date",
                "criteria": "between",
                "minimum": date(2018, 1, 1),
                "maximum": "=TODAY()",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>43101</formula1><formula2>TODAY()</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_47(self):
        """
        Test 47 Check multi range with A1 style cell ranges.
        """
        self.worksheet.data_validation(
            4,
            1,
            9,
            1,
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "multi_range": "B5:B10 D5:D10",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B10 D5:D10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_48(self):
        """
        Test 48 Check multi range with A1 style cells.
        """
        self.worksheet.data_validation(
            4,
            1,
            4,
            1,
            {
                "validate": "integer",
                "criteria": "between",
                "minimum": 1,
                "maximum": 10,
                "other_cells": [[4, 3, 4, 3]],
                "multi_range": "B5 C5",
            },
        )

        self.worksheet._write_data_validations()

        exp = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5 C5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)
