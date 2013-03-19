#
# Simple Makefile for the XlsxWriter project.
#

.PHONY: docs

docs:
	@make -C dev/docs html

pdf:
	@make -C dev/docs latexpdf

cleandocs:
	@make -C dev/docs clean

releasedocs: cleandocs docs pdf
	@cp -r dev/docs/build/html docs
	@cp -r dev/docs/build/latex/XlsxWriter.pdf docs

test:
	@python -m unittest discover

install:
	@python setup.py install
	@rm -rf build

testpythons:
	@echo "Testing with Python 2.7.2:"
	@~/.pythonbrew/pythons/Python-2.7.2/bin/python -m unittest discover
	@echo "Testing with Python 2.7.3:"
	@~/.pythonbrew/pythons/Python-2.7.3/bin/python -m unittest discover
	@echo "Testing with Python 3.2:"
	@~/.pythonbrew/pythons/Python-3.2/bin/python   -m unittest discover
	@echo "Testing with Python 3.3.0:"
	@~/.pythonbrew/pythons/Python-3.3.0/bin/python -m unittest discover

pep8:
	@ls -1 xlsxwriter/*.py | grep -v theme.py | xargs pep8
	@find xlsxwriter/test -name \*.py | xargs pep8 --ignore=E501

releasecheck:
	@dev/release/release_check.sh

release: releasecheck
	@git push origin master
	@git push --tags
	@python setup.py sdist upload
	@curl -X POST http://readthedocs.org/build/6277
	@rm -rf build

