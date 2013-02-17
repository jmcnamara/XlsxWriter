:tocdepth: 1

.. _faq:

Frequently Asked Questions
==========================

The section outlines some answers to frequently asked questions.



Q. Why do my formulas show a zero result in some, non-Excel applications?
-------------------------------------------------------------------------

Due to wide range of possible formulas and interdependencies between them
XlsxWriter doesn't, and realistically cannot, calculate the result of a
formula when it is written to an XLSX file. Instead, it stores the value 0 as
the formula result. It then sets a global flag in the XLSX file to say that
all formulas and functions should be recalculated when the file is opened.

This is the method recommended in the Excel documentation and in general it
works fine with spreadsheet applications. However, applications that don't
have a facility to calculate formulas, such as Excel Viewer, or several mobile
applications, will only display the 0 results.

If required, it is also possible to specify the calculated result of the
formula using the optional ``value`` parameter in :func:`write_formula()`::

    worksheet.write_formula('A1', '=2+2', num_format, 4)


Q. Can I apply a format to a range of cells in one go?
------------------------------------------------------

Currently no. However, it is a planned features to allow cell formats and
data to be written separately.


Q. Is feature X supported or will it be supported?
--------------------------------------------------

All supported features are documented.

Future features will match features that are available in Excel::Writer::XLSX.
Check the comparison matrix in the :ref:`ewx` section.


Q. Is there an "AutoFit" option for columns?
--------------------------------------------

Unfortunately, there is no way to specify "AutoFit" for a column
in the Excel file format. This feature is only available at runtime from
within Excel. It is possible to simulate "AutoFit" by tracking the width of
the data in the column as your write it.


Q. Do people actually ask these questions frequently, or at all?
----------------------------------------------------------------

Apart from this question, yes.

