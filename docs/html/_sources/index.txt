Create Excel files with Python and XlsxWriter
=============================================

XlsxWriter is a Python module for creating Excel XLSX files.

XlsxWriter supports the following features:

* 100% compatible Excel XLSX files.
* Write text, numbers, formulas, dates.
* Full cell formatting.
* Multiple worksheets.

Here is a small example:

.. code-block:: python

    from xlsxwriter.workbook import Workbook
        
    # Create an new Excel file and add a worksheet.
    workbook = Workbook('demo.xlsx')
    worksheet = workbook.add_worksheet()
    
    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)
    
    # Add a bold format to highlight cell text.
    bold = workbook.add_format({'bold': 1})
    
    # Write some simple text.
    worksheet.write('A1', 'Hello')
    
    # Text with formatting.
    worksheet.write('A2', 'World', bold)
    
    # Write some numbers, with row/column notation.
    worksheet.write(2, 0, 123)
    worksheet.write(3, 0, 123.456)

    workbook.close()

Which generates a worksheet like this:

.. image:: _static/intro01.png

This document explains how to install and use the XlsxWriter module.

Contents:

.. toctree::
   :maxdepth: 1
   
   introduction.rst
   getting_started.rst
   tutorial01.rst
   tutorial02.rst
   tutorial03.rst
   workbook.rst
   worksheet.rst
   format.rst
   working_with_cell_notation.rst
   working_with_formats.rst
   working_with_dates_and_time.rst
   excel_writer_xlsx.rst
   alternatives.rst
   bugs.rst
   faq.rst
   changes.rst
   author.rst
   license.rst
   
The XlsxWriter module is a port of the Perl :ref:`ewx` module. It is a work in
progress.

