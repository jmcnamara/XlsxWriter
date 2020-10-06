#
# Simple Makefile for the XlsxWriter project.
#

.PHONY: docs

docs doc:
	@make -C dev/docs html

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
	@python -m unittest discover

# Test with stable Python 2/3 releases.
testpythons:
	@echo "Testing with Python 2.7.4:"
	@~/.pythonbrew/pythons/Python-2.7.4/bin/python -m unittest discover
	@echo "Testing with Python 3.9.0:"
	@~/.pythonbrew/pythons/Python-3.9.0/bin/python -m unittest discover

# Test with all stable Python 2/3 releases.
testpythonsall:
	@echo "Testing with Python 2.7.4:"
	@~/.pythonbrew/pythons/Python-2.7.4/bin/py.test -q
	@echo "Testing with Python 3.4.1:"
	@~/.pythonbrew/pythons/Python-3.4.1/bin/py.test -q
	@echo "Testing with Python 3.5.0:"
	@~/.pythonbrew/pythons/Python-3.5.0/bin/py.test -q
	@echo "Testing with Python 3.6.6:"
	@~/.pythonbrew/pythons/Python-3.6.6/bin/py.test -q
	@echo "Testing with Python 3.7.0:"
	@~/.pythonbrew/pythons/Python-3.7.0/bin/py.test -q
	@echo "Testing with Python 3.8.0:"
	@~/.pythonbrew/pythons/Python-3.8.0/bin/py.test -q
	@echo "Testing with Python 3.9.0:"
	@~/.pythonbrew/pythons/Python-3.9.0/bin/py.test -q

test_codestyle testpep8:
	@ls -1 xlsxwriter/*.py | egrep -v "theme|compat|__init__" | xargs pycodestyle
	@pycodestyle --ignore=E501 xlsxwriter/theme.py
	@find xlsxwriter/test -name \*.py | xargs pycodestyle --ignore=E501

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
	@git push origin master
	@git push --tags
	@python setup.py sdist bdist_wheel
	@twine upload dist/*
	@../build_readthedocs.sh
	@rm -rf dist
	@rm -rf build
	@rm -rf XlsxWriter.egg-info/
