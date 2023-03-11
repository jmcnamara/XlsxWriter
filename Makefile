#
# Simple Makefile for the XlsxWriter project.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

.PHONY: docs

docs doc:
	@make -C dev/docs html

docs_external:
	@../build_readthedocs.sh

pdf:
	@make -C dev/docs latexpdf

linkcheck:
	@make -C dev/docs linkcheck

clean:
	@make -C dev/docs clean

alldocs: clean docs pdf
	@cp -r dev/docs/build/html docs
	@cp -r dev/docs/build/latex/XlsxWriter.pdf docs

pdf_release: pdf
	@cp -r dev/docs/build/latex/XlsxWriter.pdf docs

install:
	@python setup.py install
	@rm -rf build

test:
	@~/.pythonbrew/pythons/Python-3.9.0/bin/python -m unittest discover

# Test with stable Python 3 releases.
testpythons:
	@echo "Testing with Python 3.6.6:"
	@~/.pythonbrew/pythons/Python-3.6.6/bin/py.test -q
	@echo "Testing with Python 3.7.0:"
	@~/.pythonbrew/pythons/Python-3.7.0/bin/py.test -q
	@echo "Testing with Python 3.8.0:"
	@~/.pythonbrew/pythons/Python-3.8.0/bin/py.test -q
	@echo "Testing with Python 3.9.0:"
	@~/.pythonbrew/pythons/Python-3.9.0/bin/py.test -q
	@echo "Testing with Python 3.10.0:"
	@~/.pythonbrew/pythons/Python-3.10.0/bin/py.test -q
	@echo "Testing with Python 3.11.1:"
	@~/.pythonbrew/pythons/Python-3.11.1/bin/py.test -q

test_flake8:
	@ls -1 xlsxwriter/*.py | egrep -v "theme|__init__" | xargs flake8 --show-source
	@flake8 --ignore=E501 xlsxwriter/theme.py
	@find xlsxwriter/test -name \*.py | xargs flake8 --ignore=E501,F841

lint:
	@ruff xlsxwriter/*.py
	@ruff xlsxwriter/test --ignore=E501,F841
	@ruff examples

tags:
	$(Q)rm -f TAGS
	$(Q)etags xlsxwriter/*.py

testwarnings:
	@python -Werror -c 'from xlsxwriter import Workbook'

spellcheck:
	@for f in dev/docs/source/*.rst;           do aspell --lang=en_US --check $$f; done
	@for f in *.md;                            do aspell --lang=en_US --check $$f; done
	@for f in xlsxwriter/*.py;                 do aspell --lang=en_US --check $$f; done
	@for f in xlsxwriter/test/comparison/*.py; do aspell --lang=en_US --check $$f; done
	@for f in examples/*.py;                   do aspell --lang=en_US --check $$f; done
	@aspell --lang=en_US --check Changes

releasecheck:
	@dev/release/release_check.sh

release: releasecheck
	@git push origin main
	@git push --tags

	@rm -rf dist/ build/ XlsxWriter.egg-info/
	@python3 setup.py sdist bdist_wheel
	@twine upload dist/*
	@rm -rf dist/ build/ XlsxWriter.egg-info/

	@../build_readthedocs.sh
