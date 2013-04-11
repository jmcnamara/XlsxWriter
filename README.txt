XlsxWriter is a Python module for creating Excel XLSX files.

XlsxWriter supports the following features:

* 100% compatible Excel XLSX files.
* Write text, numbers, formulas, dates to cells.
* Write hyperlinks to cells.
* Full cell formatting.
* Multiple worksheets.
* Page setup methods for printing.
* Merged cells.
* Defined names.
* Autofilters.
* Data validation and drop down lists.
* Conditional formatting.
* Worksheet PNG/JPEG images.
* Rich multi-format strings.
* Cell comments.
* Document properties.
* Worksheet cell protection.
* Freeze and split worksheet panes.
* Worksheet Tables.
* Sparklines.
* Outlines and Grouping.
* Memory optimisation mode for writing large files.
* Standard libraries only.
* Python 2.6, 2.7, 3.1, 3.2 and 3.3 support.

Here is a small example::

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

See the full documentation at https://xlsxwriter.readthedocs.org/en/latest/

Release notes: https://xlsxwriter.readthedocs.org/en/latest/changes.html
   
