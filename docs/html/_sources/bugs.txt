.. _bugs:

Known Issues and Bugs
=====================

This section lists known issues and bugs and gives some information on how to
submit bug reports.

Formula results not displaying in Excel
---------------------------------------

Some early versions of Excel 2007 do not display the calculated values of
formulas written by XlsxWriter. Applying all available Service Packs to Excel
should fix this.

Formula results displaying as zero in non-Excel apps
----------------------------------------------------

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


Reporting Bugs
==============

Here are some tips on reporting bugs in XlsxWriter.


Upgrade to the latest version of the module
-------------------------------------------

The bug you are reporting may already be fixed in the latest version of the
module. Check the :ref:`changes` section as well.

Read the documentation
----------------------

The XlsxWriter documentation has been refined in response to user questions.
Therefore, if you have a question it is possible that someone else has asked
it before you and that it is already addressed in the documentation.

Look at the example programs
----------------------------

There are several example programs in the distribution. Many of these were
created in response to user questions. Try to identify an example program
that corresponds to your query and adapt it to your needs.

Use the official XlsxWriter Issue tracker on GitHub
---------------------------------------------------

The official XlsxWriter `Issue tracker is on GitHub
<https://github.com/jmcnamara/XlsxWriter/issues>`_.


Pointers for submitting a bug report
------------------------------------

1. Describe the problem as clearly and as concisely as possible.
2. Include a sample program. This is probably the most important step. Also,
   it is often easier to describe a problem in code than in written prose.
3. The sample program should be as small as possible to demonstrate the
   problem. Don't copy and past large sections of your program. The
   program should also be self contained and working.

A sample bug report is shown below. If you use this format then it will help
to analyse your question and respond to it more quickly.

   **XlsxWriter Issue with SOMETHING**

   I am using XlsxWriter and I have encountered a problem. I want it to do
   SOMETHING but the module appears to do SOMETHING ELSE.

   I am using Python version X.Y.Z and XlsxWriter x.y.z.

   Here is some code that demonstrates the problem::

       from xlsxwriter.workbook import Workbook

       workbook = Workbook('hello.xlsx')
       worksheet = workbook.add_worksheet()

       worksheet.write('A1', 'Hello world')

       workbook.close()




