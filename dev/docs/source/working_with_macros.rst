.. _macros:

Working with VBA Macros
=======================

This section explains how to add a VBA file containing functions or macros to an XlsxWriter file.

.. image:: _images/macros.png

**Note: This feature should be considered as experimental.**


The Excel XLSM file format
--------------------------

An Excel ``xlsm`` file is exactly the same as a ``xlsx`` file except that is
contains an additional ``vbaProject.bin`` file which contains functions and/or
macros. Excel uses a different extension to differentiate between the two file
formats since files containing macros are usually subject to additional
security checks.


How VBA macros are included in XlsxWriter
-----------------------------------------

The ``vbaProject.bin`` file is a binary OLE COM container. This was the format
used in older ``xls`` versions of Excel prior to Excel 2007. Unlike all of the
other components of an xlsx/xlsm file the data isn't stored in XML
format. Instead the functions and macros as stored as pre-parsed binary
format. As such it wouldn't be feasible to define macros and create a
``vbaProject.bin`` file from scratch (at least not in the remaining lifespan
and interest levels of the author).

Instead a workaround is used to extract ``vbaProject.bin`` files from existing
xlsm files and then add these to XlsxWriter files.


The vba_extract utility
-----------------------

The ``vba_extract`` utility is used to extract the ``vbaProject.bin`` binary
from an Excel 2007+ xlsm file. The utility is included in the XlsxWriter
examples directory and is also installed as a standalone executable file::

    $ vba_extract.py macro_file.xlsm
    Extracted: vbaProject.bin


Adding the VBA macros to a XlsxWriter file
------------------------------------------

Once the ``vbaProject.bin`` file has been extracted it can be added to the
XlsxWriter workbook using the :func:`add_vba_project` method::

    workbook.add_vba_project('./vbaProject.bin')

If the VBA file contains functions you can then refer to them in calculations
using :func:`write_formula`::

    worksheet.write_formula('A1', '=MyMortgageCalc(200000, 25)')

Excel files that contain functions and macros should use an ``xlsm`` extension
or else Excel will complain and possibly not open the file::

    workbook = xlsxwriter.Workbook('macros.xlsm')

It is also possible to assign a macro to a button that is inserted into a
worksheet using the :func:`insert_button` method::

    import xlsxwriter

    # Note the file extension should be .xlsm.
    workbook = xlsxwriter.Workbook('macros.xlsm')
    worksheet = workbook.add_worksheet()

    worksheet.set_column('A:A', 30)

    # Add the VBA project binary.
    workbook.add_vba_project('./vbaProject.bin')

    # Show text for the end user.
    worksheet.write('A3', 'Press the button to say hello.')

    # Add a button tied to a macro in the VBA project.
    worksheet.insert_button('B3', {'macro':   'say_hello',
                                   'caption': 'Press Me',
                                   'width':   80,
                                   'height':  30})

    workbook.close()

It may be necessary to specify a more explicit macro name prefixed by the
workbook VBA name as follows::

    worksheet.insert_button('B3', {'macro': 'ThisWorkbook.say_hello'})

See :ref:`ex_macros` from the examples directory for a working example.

.. Note::
   Button is the only VBA Control supported by Xlsxwriter. Due to the large
   effort in implementation (1+ man months) it is unlikely that any other form
   elements will be added in the future.


Setting the VBA codenames
-------------------------

VBA macros generally refer to workbook and worksheet objects. If the VBA
codenames aren't specified then XlsxWriter will use the Excel defaults of
``ThisWorkbook`` and ``Sheet1``, ``Sheet2`` etc.

If the macro uses other codenames you can set them using the workbook and
worksheet ``set_vba_name()`` methods as follows::

      # Note: set codename for workbook and any worksheets.
      workbook.set_vba_name('MyWorkbook')
      worksheet1.set_vba_name('MySheet1')
      worksheet2.set_vba_name('MySheet2')

You can find the names that are used in the VBA editor or by unzipping the
``xlsm`` file and grepping the files. The following shows how to do that using
`libxml's xmllint <http://xmlsoft.org/xmllint.html>`_ to format the XML for
clarity::


    $ unzip myfile.xlsm -d myfile
    $ xmllint --format `find myfile -name "*.xml" | xargs` | grep "Pr.*codeName"

      <workbookPr codeName="MyWorkbook" defaultThemeVersion="124226"/>
      <sheetPr codeName="MySheet"/>


.. Note::

   This step is particularly important for macros created with non-English
   versions of Excel.



What to do if it doesn't work
-----------------------------

As stated at the start of this section this feature is experimental. The
Xlsxwriter test suite contains several tests and there is a working example as
shown above. However, there is no guarantee that it will work in all
cases. Some effort may be required and some knowledge of VBA will certainly
help. If things don't work out here are some things to try:

#. Start with a simple macro file, ensure that it works and then add complexity.

#. Try to extract the macros from an Excel 2007 file. The method should work
   with macros from later versions (it was also tested with Excel 2010
   macros). However there may be features in the macro files of more recent
   version of Excel that aren't backward compatible.

#. Check the code names that macros use to refer to the workbook and
   worksheets (see the previous section above). In general VBA uses a code
   name of ``ThisWorkbook`` to refer to the current workbook and the sheet
   name (such as ``Sheet1``) to refer to the worksheets. These are the
   defaults used by XlsxWriter. If the macro uses other names then you can
   specify these using the workbook and worksheet :func:`set_vba_name`
   methods::

      # Note: set codename for workbook and any worksheets.
      workbook.set_vba_name('MyWorkbook')
      worksheet1.set_vba_name('MySheet1')
      worksheet2.set_vba_name('MySheet2')
