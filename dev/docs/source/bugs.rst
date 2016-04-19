.. _bugs:

Known Issues and Bugs
=====================

This section lists known issues and bugs and gives some information on how to
submit bug reports.

"Content is Unreadable. Open and Repair"
----------------------------------------

You may occasionally see an Excel warning when opening an XlsxWriter file
like:

   Excel could not open file.xlsx because some content is unreadable. Do you
   want to open and repair this workbook.

This ominous sounding message is Excel's default warning for any validation
error in the XML used for the components of the XLSX file.

The error message and the actual file aren't helpful in debugging issues like
this. If you do encounter this warning you should open an issue on GitHub with
a program to replicate it (see :ref:`reporting_bugs`).


"Exception caught in workbook destructor. Explicit close() may be required"
---------------------------------------------------------------------------

The following exception, or similar, can occur if the :func:`close` method
isn't used at the end of the program::

    Exception Exception: Exception('Exception caught in workbook destructor.
    Explicit close() may be required for workbook.',)
    in <bound method Workbook.__del__ of <xlsxwriter.workbook.Workbookobject
    at 0x103297d50>>

Note, it is possible that this exception will also be raised as part of
another exception that occurs during workbook destruction. In either case
ensure that there is an explicit ``workbook.close()`` in the program.


Formulas displayed as ``#NAME?`` until edited
---------------------------------------------

There are a few reasons why a formula written by XlsxWriter would generate a
``#NAME?`` error in Excel:

* Invalid formula syntax.
* Non-English function names.
* Semi-colon separators instead of commas.
* Use of Excel 2010 and later functions without a prefix.

See :ref:`working_with_formulas` and :ref:`formula_errors` for a more details
and a explanation of how to debug the issue.


Formula results displaying as zero in non-Excel applications
------------------------------------------------------------

Due to wide range of possible formulas and interdependencies between them
XlsxWriter doesn't, and realistically cannot, calculate the result of a
formula when it is written to an XLSX file. Instead, it stores the value 0 as
the formula result. It then sets a global flag in the XLSX file to say that
all formulas and functions should be recalculated when the file is opened.

This is the method recommended in the Excel documentation and in general it
works fine with spreadsheet applications. However, applications that don’t
have a facility to calculate formulas, such as Excel Viewer, or several mobile
applications, will only display the 0 results.

See :ref:`formula_result` for more details and a workaround.


Strings aren't displayed in Apple Numbers in 'constant_memory' mode
-------------------------------------------------------------------

In :func:`Workbook` ``'constant_memory'`` mode XlsxWriter uses an optimization
where cell strings aren't stored in an Excel structure call "shared strings"
and instead are written "in-line".

This is a documented Excel feature that is supported by most spreadsheet
applications. One known exception is Apple Numbers for Mac where the string
data isn't displayed.


Images not displayed correctly in Excel 2001 for Mac and non-Excel applications
-------------------------------------------------------------------------------

Images inserted into worksheets via :func:`insert_image` may not display
correctly in Excel 2011 for Mac and non-Excel applications such as OpenOffice
and LibreOffice. Specifically the images may looked stretched or squashed.

This is not specifically an XlsxWriter issue. It also occurs with files created
in Excel 2007 and Excel 2010.


Charts series created from Worksheet Tables cannot have user defined names
--------------------------------------------------------------------------

In Excel, charts created from :ref:`Worksheet Tables <tables>` have a
limitation where the data series name, if specified, must refer to a cell
within the table.

To workaround this Excel limitation you can specify a user defined name in the
table and refer to that from the chart. See :ref:`charts_from_tables`.


.. _reporting_bugs:

Reporting Bugs
==============

Here are some tips on reporting bugs in XlsxWriter.


Upgrade to the latest version of the module
-------------------------------------------

The bug you are reporting may already be fixed in the latest version of the
module. You can check which version of XlsxWriter that you are using as
follows::

    python -c 'import xlsxwriter; print(xlsxwriter.__version__)'

Check the :ref:`changes` section to see what has changed in the latest versions.


Read the documentation
----------------------

Read or search the XlsxWriter documentation to see if the issue you are
encountering is already explained.

Look at the example programs
----------------------------

There are many :ref:`examples` in the distribution. Try to identify an example
program that corresponds to your query and adapt it to use as a bug report.

Use the official XlsxWriter Issue tracker on GitHub
---------------------------------------------------

The official XlsxWriter
`Issue tracker is on GitHub <https://github.com/jmcnamara/XlsxWriter/issues>`_.


Pointers for submitting a bug report
------------------------------------

#. Describe the problem as clearly and as concisely as possible.

#. Include a sample program. This is probably the most important step. It is
   generally easier to describe a problem in code than in written prose.

#. The sample program should be as small as possible to demonstrate the
   problem. Don't copy and paste large non-relevant sections of your program.

A sample bug report is shown below. This format helps to analyze and respond to
the bug report more quickly.

   **Issue with SOMETHING**

   I am using XlsxWriter to do SOMETHING but it appears to do SOMETHING ELSE.

   I am using Python version X.Y.Z and XlsxWriter x.y.z.

   Here is some code that demonstrates the problem::

       import xlsxwriter

       workbook = xlsxwriter.Workbook('hello.xlsx')
       worksheet = workbook.add_worksheet()

       worksheet.write('A1', 'Hello world')

       workbook.close()

See also how `How to create a Minimal, Complete, and Verifiable example
<http://stackoverflow.com/help/mcve>`_ from StackOverflow.
