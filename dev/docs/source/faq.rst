.. _faq:

Frequently Asked Questions
==========================

TODO






XlsxWriter doesn't calculate the value of a formula and instead stores the
value 0 as the formula results. It then sets a global flag in the Xlsx file to
say that all formulas and functions should be recalculated when the file is
opened. This is the method recommended in the Excel documentation and in
general it works fine with spreadsheet applications. However, applications
that don't have a facility to calculate formulas, such as Excel Viewer, or
some mobile applications will only display the 0 results.

Also add to bugs.

The width corresponds to the column width value that is specified in Excel. It
is approximately equal to the length of a string in the default font of
Calibri 11. Unfortunately, there is no way to specify "AutoFit" for a column
in the Excel file format. This feature is only available at runtime from
within Excel. It is possible to simulate "AutoFit" by tracking the width of
the data in the column as your write it.




