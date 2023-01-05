.. SPDX-License-Identifier: BSD-2-Clause
   Copyright 2013-2023, John McNamara, jmcnamara@cpan.org

.. _faq:

Frequently Asked Questions
==========================

The section outlines some answers to frequently asked questions.

.. _faq_rewrite:

Q. Can XlsxWriter use an existing Excel file as a template?
-----------------------------------------------------------

No.

XlsxWriter is designed only as a file *writer*. It cannot read or modify an
existing Excel file.

.. _faq_zero_result:

Q. Why do my formulas show a zero result in some, non-Excel applications?
-------------------------------------------------------------------------

Due to a wide range of possible formulas and the interdependencies between
them XlsxWriter doesn't, and realistically cannot, calculate the result of a
formula when it is written to an XLSX file. Instead, it stores the value 0 as
the formula result. It then sets a global flag in the XLSX file to say that
all formulas and functions should be recalculated when the file is opened.

This is the method recommended in the Excel documentation and in general it
works fine with spreadsheet applications. However, applications that don't
have a facility to calculate formulas, such as Excel Viewer, or several mobile
applications, will only display the 0 results.

If required, it is also possible to specify the calculated result of the
formula using the optional ``value`` parameter in :func:`write_formula()`::

    worksheet.write_formula('A1', '=2+2', None, 4)

See also :ref:`formula_result`.

Note: **LibreOffice** doesn't recalculate Excel formulas that reference other
cells by default, in which case you will get the default XlsxWriter value
of 0. You can work around this by setting the "LibreOffice Preferences ->
LibreOffice Calc -> Formula -> Recalculation on File Load" option to "Always
recalculate" (see the LibreOffice `documentation
<https://help.libreoffice.org/6.4/en-US/text/scalc/01/06080000.html>`_). Or,
you can set a blank result in the formula, which will also force
recalculation::

    worksheet.write_formula('A1', '=Sheet1!$A$1', None, '')

.. _faq_ampersand:

Q. Why do my formulas have a @ in them?
---------------------------------------

Microsoft refers to the ``@`` in formulas as the `Implicit Intersection
Operator
<https://support.microsoft.com/en-us/office/implicit-intersection-operator-ce3be07b-0101-4450-a24e-c1c999be2b34?ui=en-us&rs=en-us&ad=us>`_.
It indicates that an input range is being reduced from multiple values to a
single value. In some cases it is just a warning indicator and doesn't affect
the calculation or result. However, in practical terms it generally means that
your formula should be written as an array formula using either
:func:`write_array_formula` or :func:`write_dynamic_array_formula`.

For more details see the :ref:`formula_dynamic_arrays` and
:ref:`formula_intersection_operator` sections of the XlsxWriter documentation.

.. _faq_format_range:

Q. Can I apply a format to a range of cells in one go?
------------------------------------------------------

Currently no. However, it is a planned features to allow cell formats and data
to be written separately.

.. _faq_future:

Q. Is feature X supported or will it be supported?
--------------------------------------------------

All supported features are documented. Future features are on the `Roadmap
<https://github.com/jmcnamara/XlsxWriter/issues/653>`_.

.. _faq_protect_workbook:

Q. Can I password protect an XlsxWriter xlsx file
-------------------------------------------------

Although it is possible to password protect a worksheet using the Worksheet
:func:`protect` method it isn't possible to password protect the entire
workbook/file using XlsxWriter.

The reason for this is that a protected/encrypted xlsx file is in a different
format from an ordinary xlsx file. This would require a lot of additional work,
and testing, and isn't something that is on the XlsxWriter roadmap.

However, it is possible to password protect an XlsxWriter generated file using
a third party open source tool called `msoffice-crypt
<https://github.com/herumi/msoffice>`_. This works for macOS, Linux and Windows::

    msoffice-crypt.exe -e -p password clear.xlsx encrypted.xlsx

.. _faq_faq:

Q. Do people actually ask these questions frequently, or at all?
----------------------------------------------------------------

Apart from this question, yes.
