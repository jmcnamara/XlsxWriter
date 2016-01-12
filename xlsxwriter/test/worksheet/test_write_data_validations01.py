###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from datetime import date
from ...compatibility import StringIO
from ...worksheet import Worksheet
from ..helperfunctions import _xml_to_list


class TestWriteDataValidations(unittest.TestCase):
    """
    Test the Worksheet _write_data_validations() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_data_validations_1(self):
        """Test the _write_data_validations() method. Data validation example 1 from docs"""

        self.worksheet.data_validation('A1', {'validate': 'integer',
                                              'criteria': '>',
                                              'value': 0,
                                              })

        self.worksheet._write_data_validations()

        exp = """<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>"""
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_1b(self):
        """Test the _write_data_validations() method. Data validation example 1 from docs (with options turned off)"""

        self.worksheet.data_validation('A1', {'validate': 'integer',
                                              'criteria': '>',
                                              'value': 0,
                                              'ignore_blank': 0,
                                              'show_input': 0,
                                              'show_error': 0,
                                              })

        self.worksheet._write_data_validations()

        exp = """<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>"""
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_2(self):
        """Test the _write_data_validations() method. Data validation example 2 from docs"""

        self.worksheet.data_validation('A2', {'validate': 'integer',
                                              'criteria': '>',
                                              'value': '=E3',
                                              })

        self.worksheet._write_data_validations()

        exp = """<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A2"><formula1>E3</formula1></dataValidation></dataValidations>"""
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_3(self):
        """Test the _write_data_validations() method. Data validation example 3 from docs"""

        self.worksheet.data_validation('A3', {'validate': 'decimal',
                                              'criteria': 'between',
                                              'minimum': 0.1,
                                              'maximum': 0.5,
                                              })

        self.worksheet._write_data_validations()

        exp = """<dataValidations count="1"><dataValidation type="decimal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A3"><formula1>0.1</formula1><formula2>0.5</formula2></dataValidation></dataValidations>"""
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_4(self):
        """Test the _write_data_validations() method. Data validation example 4 from docs"""

        self.worksheet.data_validation('A4', {'validate': 'list',
                                              'source': ['open', 'high', 'close'],
                                              })

        self.worksheet._write_data_validations()

        exp = """<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A4"><formula1>"open,high,close"</formula1></dataValidation></dataValidations>"""
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_5(self):
        """Test the _write_data_validations() method. Data validation example 5 from docs"""

        self.worksheet.data_validation('A5', {'validate': 'list',
                                              'source': '=$E$4:$G$4',
                                              })

        self.worksheet._write_data_validations()

        exp = """<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A5"><formula1>$E$4:$G$4</formula1></dataValidation></dataValidations>"""
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_6(self):
        """Test the _write_data_validations() method. Data validation example 6 from docs"""

        self.worksheet.data_validation('A6', {'validate': 'date',
                                              'criteria': 'between',
                                              'minimum': date(2008, 1, 1),
                                              'maximum': date(2008, 12, 12),
                                              })

        self.worksheet._write_data_validations()

        exp = """<dataValidations count="1"><dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A6"><formula1>39448</formula1><formula2>39794</formula2></dataValidation></dataValidations>"""
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)

    def test_write_data_validations_7(self):
        """Test the _write_data_validations() method. Data validation example 7 from docs"""

        self.worksheet.data_validation('A7', {'validate': 'integer',
                                              'criteria': 'between',
                                              'minimum': 1,
                                              'maximum': 100,
                                              'input_title': 'Enter an integer:',
                                              'input_message': 'between 1 and 100',
                                              })

        self.worksheet._write_data_validations()

        exp = """<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Enter an integer:" prompt="between 1 and 100" sqref="A7"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>"""
        got = self.fh.getvalue()

        exp = _xml_to_list(exp)
        got = _xml_to_list(got)

        self.assertEqual(got, exp)
