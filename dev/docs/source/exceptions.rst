.. SPDX-License-Identifier: BSD-2-Clause
   Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org

.. _exceptions:

The Exceptions Class
====================

The Exception class contains the various exceptions that can be raised by
XlsxWriter. In general XlsxWriter only raised exceptions for un-recoverable
errors or for errors that would lead to file corruption such as creating two
worksheets with the same name.

The hierarchy of exceptions in XlsxWriter is:

* ``XlsxWriterException(Exception)``

  * ``XlsxFileError(XlsxWriterException)``

    * ``FileCreateError(XlsxFileError)``

    * ``UndefinedImageSize(XlsxFileError)``

    * ``UndefinedImageSize(XlsxFileError)``

    * ``FileSizeError(XlsxFileError)``

  * ``XlsxInputError(XlsxWriterException)``

    * ``DuplicateTableName(XlsxInputError)``

    * ``InvalidWorksheetName(XlsxInputError)``

    * ``DuplicateWorksheetName(XlsxInputError)``

    * ``OverlappingRange(XlsxInputError)``

    * ``ThemeFileError(XlsxInputError)``


Exception: XlsxWriterException
------------------------------

.. py:exception:: XlsxWriterException


Base exception for XlsxWriter.


Exception: XlsxFileError
------------------------

.. py:exception:: XlsxFileError


Base exception for all file related errors.


Exception: XlsxInputError
-------------------------

.. py:exception:: XlsxInputError


Base exception for all input data related errors.


Exception: FileCreateError
--------------------------

.. py:exception:: FileCreateError

This exception is raised if there is a file permission, or IO error, when
writing the xlsx file to disk. This can be caused by an non-existent directory
or (in Windows) if the file is already open in Excel::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    worksheet = workbook.add_worksheet()

    # The file exception.xlsx is already open in Excel.
    workbook.close()

Raises::

    xlsxwriter.exceptions.FileCreateError:
        [Errno 13] Permission denied: 'exception.xlsx'

This exception can be caught in a ``try`` block where you can instruct the
user to close the open file before overwriting it::

    while True:
        try:
            workbook.close()
        except xlsxwriter.exceptions.FileCreateError as e:
            decision = input("Exception caught in workbook.close(): %s\n"
                             "Please close the file if it is open in Excel.\n"
                             "Try to write file again? [Y/n]: " % e)
            if decision != 'n':
                continue

        break

See also :ref:`ex_check_close`.


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
one of the supported file formats: PNG, JPEG, GIF, BMP, WMF or EMF. The exception
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


Exception: FileSizeError
------------------------

.. py:exception:: FileSizeError

This exception is raised if one of the XML files that is part of the xlsx file, or the xlsx file itself, exceeds 4GB in size::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    worksheet = workbook.add_worksheet()

    # Write lots of data to create a very big file.

    workbook.close()

Raises::

    xlsxwriter.exceptions.FileSizeError:
        Filesize would require ZIP64 extensions. Use workbook.use_zip64().

As noted in the exception message, files larger than 4GB can be created by
turning on the zipfile.py ZIP64 extensions using the :func:`use_zip64` method.



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

Or for a worksheet name start or ends with an apostrophe::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    worksheet = workbook.add_worksheet("'Sheet1'")

    workbook.close()

Raises::

    xlsxwriter.exceptions.InvalidWorksheetName:
        Sheet name cannot start or end with an apostrophe "'Sheet1'".


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


Exception: OverlappingRange
---------------------------------

.. py:exception:: OverlappingRange

This exception is raised during Worksheet :func:`add_table()` or
:func:`merge_range()` if the range overlaps an existing worksheet table or merge
range. This is a file corruption error in Excel::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    worksheet = workbook.add_worksheet()

    worksheet.merge_range('A1:G10', 'Range 1')
    worksheet.merge_range('G10:K20', 'Range 2')

    workbook.close()

Raises::

    xlsxwriter.exceptions.OverlappingRange:
        Merge range 'G10:K20' overlaps previous merge range 'A1:G10'.

Exception: ThemeFileError
-------------------------

.. py:exception:: ThemeFileError

This exception is raised during Workbook :func:`use_custom_theme()` if the theme
file is invalid or contains unsupported elements such as image fills::

    import xlsxwriter

    workbook = xlsxwriter.Workbook('exception.xlsx')

    workbook.use_custom_theme("theme.xml")

    worksheet = workbook.add_worksheet()

    workbook.close()

Raises::

    xlsxwriter.exceptions.ThemeFileError:
        Invalid XML theme file: 'theme.xml'.