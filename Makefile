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
	@echo "Testing with Python 3.4.1:"
	@~/.pythonbrew/pythons/Python-3.4.1/bin/python -m unittest discover

# Test with all stable Python 2/3 releases.
testpythonsall:
	@echo "Testing with Python 2.5.6:"
	@~/.pythonbrew/pythons/Python-2.5.6/bin/py.test -q
	@echo "Testing with Python 2.6.8:"
	@~/.pythonbrew/pythons/Python-2.6.8/bin/py.test -q
	@echo "Testing with Python 2.7.4:"
	@~/.pythonbrew/pythons/Python-2.7.4/bin/py.test -q
	@echo "Testing with Python 3.1.5:"
	@~/.pythonbrew/pythons/Python-3.1.5/bin/py.test -q
	@echo "Testing with Python 3.2.5:"
	@~/.pythonbrew/pythons/Python-3.2.5/bin/py.test -q
	@echo "Testing with Python 3.3.2:"
	@~/.pythonbrew/pythons/Python-3.3.2/bin/py.test -q
	@echo "Testing with Python 3.4.1:"
	@~/.pythonbrew/pythons/Python-3.4.1/bin/py.test -q
	@echo "Testing with Python 3.5.0:"
	@~/.pythonbrew/pythons/Python-3.5.0/bin/py.test -q

testpep8:
	@ls -1 xlsxwriter/*.py | egrep -v "theme|compat|__init__" | xargs flake8
	@pep8 --ignore=E501 xlsxwriter/theme.py
	@pep8 --ignore=E501 xlsxwriter/compat_collections.py
	@find xlsxwriter/test -name \*.py | xargs pep8 --ignore=E501

spellcheck:
	@for f in dev/docs/source/*.rst; do aspell --lang=en_US --check $$f; done
	@for f in *.md;                  do aspell --lang=en_US --check $$f; done
	@for f in xlsxwriter/*.py;       do aspell --lang=en_US --check $$f; done
	@for f in examples/*.py;         do aspell --lang=en_US --check $$f; done
	@aspell --lang=en_US --check Changes

releasecheck:
	@dev/release/release_check.sh

release: releasecheck
	@git push origin master
	@git push --tags
	@python setup.py sdist bdist_wheel
	@twine upload dist/*
	@curl -X POST https://readthedocs.org/build/xlsxwriter/latest
	@rm -rf dist
	@rm -rf build
	@rm -rf XlsxWriter.egg-info/
