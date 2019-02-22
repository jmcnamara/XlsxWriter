XlsxWriter
==========

**XlsxWriter** is a Python module for writing files in the Excel 2007+ XLSX
file format.

XlsxWriter can be used to write text, numbers, formulas and hyperlinks to
multiple worksheets and it supports features such as formatting and many more,
including:

* 100% compatible Excel XLSX files.
* Full formatting.
* Merged cells.
* Defined names.
* Charts.
* Autofilters.
* Data validation and drop down lists.
* Conditional formatting.
* Worksheet PNG/JPEG/BMP/WMF/EMF images.
* Rich multi-format strings.
* Cell comments.
* Integration with Pandas.
* Textboxes.
* Memory optimization mode for writing large files.

It supports Python 2.7, 3.4+, Jython and PyPy and uses standard libraries only.

Here is a simple example:

.. code-block:: python

   import xlsxwriter


   # Create an new Excel file and add a worksheet.
   workbook = xlsxwriter.Workbook('demo.xlsx')
   worksheet = workbook.add_worksheet()

   # Widen the first column to make the text clearer.
   worksheet.set_column('A:A', 20)

   # Add a bold format to use to highlight cells.
   bold = workbook.add_format({'bold': True})

   # Write some simple text.
   worksheet.write('A1', 'Hello')

   # Text with formatting.
   worksheet.write('A2', 'World', bold)

   # Write some numbers, with row/column notation.
   worksheet.write(2, 0, 123)
   worksheet.write(3, 0, 123.456)

   # Insert an image.
   worksheet.insert_image('B5', 'logo.png')

   workbook.close()

.. image:: https://raw.github.com/jmcnamara/XlsxWriter/master/dev/docs/source/_images/demo.png

See the full documentation at: https://xlsxwriter.readthedocs.io

Release notes: https://xlsxwriter.readthedocs.io/changes.html

