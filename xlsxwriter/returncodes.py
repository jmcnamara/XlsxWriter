###############################################################################
#
# ReturnCodes - A class for XlsxWriter return codes.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2022, John McNamara, jmcnamara@cpan.org
#

from enum import Enum


class ReturnCode(str, Enum):
    """Return codes enumeration for XlsxWriter functions

    These values can be converted to string in different ways:
    - direct assignment: mystr = retcode
    - through format function: mystr = "Value {0}".format(retcode)
    - through f-string: mystr = f'Value {retcode}'

    Conversion through str() function will result in the internal Enum
    representation, i.e. str(ReturnCode.XW_NO_ERROR) will return
    "ReturnCode.XW_NO_ERROR"
    """

    # Note: the following are not tuples, but strings on multiple lines
    # This is required to be compliant to E501

    ###########################################################################
    #
    # Values from libxlsxwriter library
    #
    ###########################################################################

    XW_NO_ERROR = "No error"
    XW_ERROR_MAX_STRING_LENGTH_EXCEEDED = ("String exceeds Excel's limit of "
                                           "32,767 characters")
    XW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE = ("Worksheet row or column index "
                                             "out of range")
    XW_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED = ("Maximum hyperlink length "
                                                  "(2079) exceeded")
    XW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED = ("Maximum number of "
                                                   "worksheet URLs (65530) "
                                                   "exceeded")

    ###########################################################################
    #
    # Values added only for this library
    #
    ###########################################################################

    XW_ERROR_VBA_FILE_NOT_FOUND = "VBA project binary file not found"
    XW_ERROR_FORMULA_CANT_BE_NONE_OR_EMPTY = "Formula can't be None or empty"
    XW_ERROR_2_CONSECUTIVE_FORMATS = "2 consecutive formats used"
    XW_ERROR_EMPTY_STRING_USED = "Empty string used"
    XW_ERROR_INSUFFICIENT_PARAMETERS = "Insufficient parameters"
    XW_ERROR_IMAGE_FILE_NOT_FOUND = "Image file not found"
    XW_ERROR_INCORRECT_PARAMETER_OR_OPTION = "Incorrect parameter or option"
    XW_ERROR_NOT_SUPPORTED_COSTANT_MEMORY = ("Not supported in "
                                             "constant_memory mode")
