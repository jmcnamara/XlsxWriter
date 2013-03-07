.. _intro:

Introduction
============

**XlsxWriter** is a Python module for writing files in the Excel 2007+ XLSX
file format.

Multiple worksheets can be added to a workbook and formatting can be applied to
cells. Text, numbers, and formulas can be written to the cells.

This module cannot be used to modify or write to an existing Excel XLSX file.
Modifying Excel files is not, and never was, part of the design scope. There
are some :ref:`alternatives` that do that.

The XlsxWriter module is a port of the Perl ``Excel::Writer::XLSX`` module.
It is a work in progress. See the :ref:`ewx` section for a list of
currently ported features.

XlsxWriter is written by John McNamara who also wrote the perl modules
`Excel::Writer::XLSX <http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX/>`_
and
`Spreadsheet::WriteExcel <http://search.cpan.org/~jmcnamara/Spreadsheet-WriteExcel/>`_
and who is the maintainer of
`Spreadsheet::ParseExcel <http://search.cpan.org/~jmcnamara/Spreadsheet-ParseExcel/>`_.

XlsxWriter is intended to have a high degree of compatibility with files
produced by Excel. In most cases the files produced are 100% equivalent to
files produced by Excel. In fact the
`test suite <https://github.com/jmcnamara/XlsxWriter/tree/master/xlsxwriter/test/comparison>`_
contains a range of test cases that verify the output of XlsxWriter against
actual files created in Excel.

XlsxWriter is licensed under a BSD :ref:`License` and is available as a ``git``
repository on `GitHub <http://github.com/jmcnamara/XlsxWriter>`_.

