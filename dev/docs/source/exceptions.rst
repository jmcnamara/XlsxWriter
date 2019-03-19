.. _exceptions:

The Exceptions Class
====================

The Exception class contains the various exceptions that can be raised by
XlsxWriter. In general XlsxWriter only raised exceptions for un-recoverable
errors or for errors that would lead to file corruption such as creating two
worksheets with the same name.

The hierarchy of exceptions in XlsxWriter is:

* ``XlsxWriterException(Exception)``

  * ``XlsxInputError(XlsxWriterException)``

    * ``DuplicateTableName(XlsxInputError)``

    * ``InvalidWorksheetName(XlsxInputError)``

    * ``DuplicateWorksheetName(XlsxInputError)``

  * ``XlsxFileError(XlsxWriterException)``

    * ``UndefinedImageSize(XlsxFileError)``

    * ``UnsupportedImageFormat(XlsxFileError)``


Exception: XlsxWriterException
------------------------------

.. py:exception:: XlsxWriterException


Base exception for XlsxWriter.


Exception: XlsxInputError
-------------------------

.. py:exception:: XlsxInputError


Base exception for all input data related errors.


Exception: XlsxFileError
------------------------

.. py:exception:: XlsxFileError


Base exception for all file related errors.


Exception: EmptyChartSeries
---------------------------

.. py:exception:: EmptyChartSeries

This exception is raised if a chart is added to a worksheet without a data
series. The exception is raised during Workbook :func:`close()`::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')
    worksheet = workbook.add_worksheet()

    chart = workbook.add_chart({'type': 'column'})

    worksheet.insert_chart('A7', chart)

    workbook.close()

Raises::

    xlsxwriter.exceptions.EmptyChartSeries:
        Chart1 must contain at least one data series. See chart.add_series().


Exception: DuplicateTableName
-----------------------------

.. py:exception:: DuplicateTableName

This exception is raised if a duplicate worksheet table name in used via
:func:`add_table()`. The exception is raised during Workbook :func:`close()`::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.add_table('B1:F3', {'name': 'SalesData'})
    worksheet.add_table('B4:F7', {'name': 'SalesData'})

    workbook.close()

Raises::

    xlsxwriter.exceptions.DuplicateTableName:
        Duplicate name 'SalesData' used in worksheet.add_table().


Exception: InvalidWorksheetName
-------------------------------

.. py:exception:: InvalidWorksheetName

This exception is raised during Workbook :func:`add_worksheet()` if a
worksheet name is too long or contains restricted characters.

For example with a 32 character worksheet name::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    name = 'name_that_is_longer_than_thirty_one_characters'
    worksheet = workbook.add_worksheet(name)

    workbook.close()

Raises::

    xlsxwriter.exceptions.InvalidWorksheetName:
        Excel worksheet name 'name_that_is_longer_than_thirty_one_characters'
        must be <= 31 chars.

Or for a worksheet name containing one of the Excel restricted characters,
i.e. ``[ ] : * ? / \``::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    worksheet = workbook.add_worksheet('Data[Jan]')

    workbook.close()

Raises::

    xlsxwriter.exceptions.InvalidWorksheetName:
        Invalid Excel character '[]:*?/\' in sheetname 'Data[Jan]'.


Exception: DuplicateWorksheetName
---------------------------------

.. py:exception:: DuplicateWorksheetName

This exception is raised during Workbook :func:`add_worksheet()` if a
worksheet name has already been used. As with Excel the check is case
insensitive::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    worksheet1 = workbook.add_worksheet('Sheet1')
    worksheet2 = workbook.add_worksheet('sheet1')

    workbook.close()

Raises::

    xlsxwriter.exceptions.DuplicateWorksheetName:
        Sheetname 'sheet1', with case ignored, is already in use.


Exception: UndefinedImageSize
-----------------------------

.. py:exception:: UndefinedImageSize

This exception is raised if an image added via :func:`insert_image()` doesn't
contain height or width information. The exception is raised during Workbook
:func:`close()`::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    worksheet = workbook.add_worksheet()

    worksheet.insert_image('A1', 'logo.png')

    workbook.close()

Raises::

    xlsxwriter.exceptions.UndefinedImageSize:
         logo.png: no size data found in image file.

.. note::

   This is a relatively rare error that is most commonly caused by XlsxWriter
   failing to parse the dimensions of the image rather than the image not
   containing the information. In these cases you should raise a GitHub issue
   with the image attached, or provided via a link.


Exception: UnsupportedImageFormat
---------------------------------

.. py:exception:: UnsupportedImageFormat

This exception is raised if if an image added via :func:`insert_image()` isn't
one of the supported file formats: PNG, JPEG, BMP, WMF or EMF. The exception
is raised during Workbook :func:`close()`::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    worksheet = workbook.add_worksheet()

    worksheet.insert_image('A1', 'logo.xyz')

    workbook.close()

Raises::

    xlsxwriter.exceptions.UnsupportedImageFormat:
        logo.xyz: Unknown or unsupported image file format.

.. note::

   If the image type is one of the supported types, and you are sure that the
   file format is correct, then the exception may be caused by XlsxWriter
   failing to parse the type of the image correctly. In these cases you should
   raise a GitHub issue with the image attached, or provided via a link.
