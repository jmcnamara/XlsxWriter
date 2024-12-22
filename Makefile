#
# Simple Makefile for the XlsxWriter project.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2024, John McNamara, jmcnamara@cpan.org
#

.PHONY: docs

docs doc:
	@make -C dev/docs html
	@open dev/docs/build/html/index.html

docs_external:
	@../build_readthedocs.sh

linkcheck:
	@make -C dev/docs linkcheck

clean:
	@make -C dev/docs clean

install:
	@python setup.py install
	@rm -rf build

test:
	@~/.pythonbrew/pythons/Python-3.12.7/bin/python3.12 -m unittest discover

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
	@echo "Testing with Python 3.12.7:"
	@~/.pythonbrew/pythons/Python-3.12.7/bin/py.test -q
	@echo "Testing with Python 3.13.0:"
	@~/.pythonbrew/pythons/Python-3.13.0/bin/py.test -q

test_flake8:
	@ls -1 xlsxwriter/*.py | egrep -v "theme|__init__" | xargs flake8 --show-source --max-line-length=88 --ignore=E203,W503
	@flake8 --ignore=E501 xlsxwriter/theme.py
	@find xlsxwriter/test -name \*.py | xargs flake8 --ignore=E501,F841,W503

lint:
	@ruff check xlsxwriter/*.py
	@ruff check xlsxwriter/test --ignore=E501,F841
	@ruff check examples
	@black --check xlsxwriter/ examples/
	@pylint xlsxwriter/*.py
	@pylint --rcfile=examples/.pylintrc examples/*.py
	@isort --check --diff --profile black xlsxwriter/ examples/

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
