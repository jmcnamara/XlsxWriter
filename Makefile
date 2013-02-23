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

installdocs: cleandocs docs pdf
	@cp -r dev/docs/build/html docs
	@cp -r dev/docs/build/latex/XlsxWriter.pdf docs

test:
	@python -m unittest discover

testpythons:
	@pythonbrew switch 2.7.2
	@python -m unittest discover
	@pythonbrew switch 2.7.3
	@python -m unittest discover
	@pythonbrew switch 3.2
	@python -m unittest discover
	@pythonbrew switch 3.3.0
	@python -m unittest discover
	@pythonbrew switch 2.7.2

pep8:
	@ls -1 xlsxwriter/*.py | grep -v theme.py | xargs pep8
	@find xlsxwriter/test/ -name \*.py | xargs pep8 --ignore=E501

releasecheck:
	@dev/release/release_check.sh

release: releasecheck
	@git push origin master
	@git push --tags
	@python setup.py sdist upload

