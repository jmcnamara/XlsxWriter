# -*- coding: utf-8 -*-

"""
xlsxwriter.exceptions

~~~~~~~~~~~~~~~~~~~~~


This module contains the set of XlsxWriter's Exceptions
"""


class XlsxWriterException(Exception):
    """Base Exception for XlsxWriter"""


class XlsxDataError(XlsxWriterException):
    """Base Exception for all data related errors"""


class XlsxFormatError(XlsxWriterException):
    """Base Exception for all format errors"""


class XlsxFileError(XlsxWriterException):
    """Base Exception for all file/image related errors"""


class UndefinedImageSize(XlsxFileError):
    """No size data found in image file"""


class UnsupportedImageFormat(XlsxFileError):
    """Unsupported image file format"""


class WorkbookDestructorError(XlsxFileError):
    """Unable to close workbook"""


class EmptyChartSeries(XlsxDataError):
    """Chart must contain atleast one data series"""


class DuplicateTableName(XlsxDataError):
    """Table with that name already exists"""


class InvalidWorksheetName(XlsxDataError):
    """
    Worksheet name either is empty,
    too long or contains restricted characters
    """


class DuplicateWorksheetName(InvalidWorksheetName):
    """Worksheet with that name already exists"""
