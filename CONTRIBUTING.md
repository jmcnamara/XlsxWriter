# XlsxWriter: Bug Reports and Pull Requests


## Reporting Bugs

Here are some tips on reporting bugs in XlsxWriter.


### Upgrade to the latest version of the module

The bug you are reporting may already be fixed in the latest version of the
module. You can check which version of xlsxwriter that you are using as follows:

    python -c 'import xlsxwriter; print(xlsxwriter.__version__)'


The [Changes](https://github.com/jmcnamara/XlsxWriter/blob/master/Changes) file lists what has changed in the latest versions.


### Read the documentation

Read or search the [XlsxWriter documentation](https://xlsxwriter.readthedocs.io/) to see if the issue you are encountering is already explained.

### Look at the example programs

There are [many example programs](https://github.com/jmcnamara/XlsxWriter/tree/master/examples) in the distribution. Try to identify an example program that corresponds to your query and adapt it to use as a bug report.


### Pointers for submitting a bug report

1. Describe the problem as clearly and as concisely as possible.
2. Include a sample program. This is probably the most important step. It is generally easier to describe a problem in code than in written prose.
3. The sample program should be as small as possible to demonstrate the problem. Don't copy and paste large non-relevant sections of your program.


### Sample Bug Report

A sample bug report is shown below. This format helps to analyze and respond to the bug report more quickly.

> **Issue with SOMETHING**
>
>   I am using XlsxWriter to do SOMETHING but it appears to do SOMETHING ELSE.
>
>   I am using Python version X.Y.Z and XlsxWriter x.y.z.
>
>   Here is some code that demonstrates the problem::
>
>       import xlsxwriter
>
>       workbook = xlsxwriter.Workbook('hello.xlsx')
>       worksheet = workbook.add_worksheet()
>
>       worksheet.write('A1', 'Hello world')
>
>       workbook.close()


See also how [How to create a Minimal, Complete, and Verifiable example](http://stackoverflow.com/help/mcve) from StackOverflow.


### Use the XlsxWriter GitHub issue tracker

Submit the bug report using the [XlsxWriter issue tracker](https://github.com/jmcnamara/XlsxWriter/issues).


## Pull Requests and Contributing to XlsxWriter

All patches and pull requests are welcome but must start with an issue tracker.


### Getting Started

1. Pull requests and new feature proposals must start with an [issue tracker](https://github.com/jmcnamara/XlsxWriter/issues). This serves as the focal point for the design discussion.
2. Describe what you plan to do. If there are API changes or additions add some pseudo-code to demonstrate them.
3. Fork the repository.
4. Run all the tests to make sure the current code work on your system using `make test`.
5. Create a feature branch for your new feature.
6. Enable Travis CI on your fork, see below.


### Enabling Travis CI via your GitHub account

Travis CI is a free Continuous Integration service that will test any code you push to GitHub with various versions of Python 2 and 3, and PyPy.

See the [Travis CI Getting Started](http://about.travis-ci.org/docs/user/getting-started/) instructions.

Note there is already a `.travis.yml` file in the XlsxWriter repo so that doesn't need to be created.


### Writing Tests

This is the most important step. XlsxWriter has over 1000 tests and a 2:1 test to code ratio. Patches and pull requests for anything other than minor fixes or typos will not be merged without tests.

Use the existing tests in `XlsxWriter/xlsxwriter/test/` as examples.

Ideally, new features should be accompanied by tests that compare XlsxWriter output against actual Excel 2007 files. See the `XlsxWriter/xlsxwriter/test/comparison` test files for examples. If you don't have access to Excel 2007 I can help you create input files for test cases.

Tests should use the standard [unittest](http://docs.python.org/2/library/unittest.html) Python module.


### Code Style

Follow the general style of the surrounding code and format it to the [PEP8](http://www.python.org/dev/peps/pep-0008/) coding standards.

Tests should conform to `PEP8` but can ignore `E501` for long lines to allow the inclusion of Excel XML in tests.

There is a make target that will verify the source and test files:

    make testpep8


### Running tests

As a minimum, tests should be run using Python 2.7 and Python 3.5.


    make test
    # or
    py.test

I use [pythonbrew](https://github.com/utahta/pythonbrew) and [Tox](https://tox.readthedocs.io/en/latest/) to test with a variety of Python versions. See the Makefile for example test targets. A `tox.ini` file is already configured.

When you push your changes they will also be tested using [Travis CI](https://travis-ci.org/jmcnamara/XlsxWriter/) as explained above.


### Documentation

If your feature requires it then write some [RST](http://docutils.sourceforge.net/rst.html) documentation in [Sphinx](http://sphinx-doc.org) format or add to the existing documentation.

The docs, in `dev/docs/source` can be built in Html format using:

    make docs


### Example programs

If applicable add an example program to the `examples` directory.


### Copyright and License

Copyright remains with the original author. Do not include additional copyright claims or Licensing requirements. GitHub and the `git` repository will record your contribution an it will be acknowledged in the Changes file.


### Submitting the Pull Request

If your change involves several incremental `git` commits then `rebase` or `squash` them onto another branch so that the Pull Request is a single commit or a small number of logical commits.

Push your changes to GitHub and submit the Pull Request with a hash link to the to the Issue tracker that was opened above.
