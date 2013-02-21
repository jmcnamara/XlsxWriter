#
# Simple Makefile for the XlsxWriter project.
#

.PHONY: docs cleandocs installdocs test sdist pep8

docs:
	@make -C dev/docs html

cleandocs:
	@make -C dev/docs clean

installdocs: cleandocs docs
	@cp -r dev/docs/build/html docs

test:
	@python -m unittest discover

sdist:
	@python setup.py sdist

pep8:
	@ls -1 xlsxwriter/*.py | grep -v theme.py | xargs pep8
	@find xlsxwriter/test/ -name \*.py | xargs pep8 --ignore=E501
