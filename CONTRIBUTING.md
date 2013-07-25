## How to Contribute to XlsxWriter

All patches and pull requests should be made via [Github](https://github.com/jmcnamara/XlsxWriter).


## Getting Started

1. For bug fixes submit an issue in the XlsxWriter [issue tracker](https://github.com/jmcnamara/XlsxWriter/issues). See the [suggestions on submitting a bug report](https://github.com/jmcnamara/XlsxWriter/issues/1).
2. New features should also start with an issue tracker. Describe what you plan to do. If possible try to gauge if there is general interest in the feature that you are proposing.
3. Fork the repository.
4. Run all the tests to make sure the code works using `make test`.


## Write Tests

This is the most important step. XlsxWriter has approximately 1000 tests and a 2:1 test to code ratio. Patches and pull requests for anything other than minor fixes or typos will not be accepted without tests.

Use the existing tests in `XlsxWriter/xlsxwriter/test/` as examples.

New features should be accompanied by tests that compare XlsxWriter output against actual Excel 2007 files. See the `XlsxWriter/xlsxwriter/test/comparison` test files for examples. Tests against other versions of Excel or other Spreadsheet applications aren't appropriate.

Tests should use the standard [unittest](http://docs.python.org/2/library/unittest.html) Python module.


## Write Code

Follow the general style of the surrounding code and format it to the [PEP8](http://www.python.org/dev/peps/pep-0008/) coding standards.

Tests should conform to `PEP8` but can ignore `E501` for long lines to allow the inclusion of Excel XML in tests.

There is a make target that will verify the source and test files:

    make pep8


## Run the tests

Tests should be run using Python 2 and Python 3.

The author uses [pythonbrew](https://github.com/utahta/pythonbrew) to test with a variety of Python versions. See the Makefile for example test targets.

When you push your changes they will also be tested for a variety of Python versions using [Travis CI](https://travis-ci.org/jmcnamara/XlsxWriter/).


## Write Documentation

Write some [rST](http://docutils.sourceforge.net/rst.html) documentation in [Sphinx](http://sphinx-doc.org) format or add to the existing documentation.

The docs can be built using:

    make docs


## Write an example program

If applicable add an example program to the `examples` directory.


## Copyright and License

Copyright remains with the original author. Do not include additional copyright claims or Licensing requirements.


## Submit the changes

Push your changes to GitHub and submit a Pull Request. Ideally that should be attached to the Issue tracker that was opened above.


## Too much work?

Best effort counts as well.

And remember, the module author went though each of these steps for each (30+) release of the module.
