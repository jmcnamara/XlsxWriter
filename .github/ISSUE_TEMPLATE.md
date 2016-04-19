# Opening XlsxWriter Issues

## Asking questions about using XlsxWriter

General questions on how to do something with the XlsxWriter module should be
asked on [StackOverflow](http://stackoverflow.com/questions/tagged/xlsxwriter).
Add the ``xlsxwriter`` tag to the question. This has a better chance of
getting several answers and also helps others who might have similar questions
in the future.


See below for information on adding a Bug Report or Feature Request.


## Reporting Bugs

Modify the following to suit.


Title: Issue with SOMETHING

Hi,

I am using XlsxWriter to do SOMETHING but it appears to do SOMETHING ELSE.

I am using Python version X.Y.Z and XlsxWriter x.y.z and Excel version X.

Here is some code that demonstrates the problem:

```python

import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Hello world')

workbook.close()

```

Notes:

1. Ensure that the example code can be run to generate a file that
   demonstrates the issue.
2. Only include code that relates to the issue.
3. Remove non-relevant text from this template.
4. If you are seeing an issue in LibreOffice, OpenOffice or another third
   party application, also test the output with a version of Excel.
5. You can get the required version numbers as follows:

        python --version
        python -c 'import xlsxwriter; print(xlsxwriter.__version__)'


## Feature requests

Add `Feature request:` to the title.

In the comment section describe the feature that you would like to be added.

If you are currently using a workaround you can show that workaround.
