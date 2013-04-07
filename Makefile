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

alldocs: cleandocs docs pdf
	@cp -r dev/docs/build/html docs
	@cp -r dev/docs/build/latex/XlsxWriter.pdf docs

install:
	@python setup.py install
	@rm -rf build

test:
	@python -m unittest discover

testpythons:
    # Test with Python 2/3 stable releases.
	@echo "Testing with Python 2.7.4:"
	@~/.pythonbrew/pythons/Python-2.7.4/bin/python -m unittest discover
	@echo "Testing with Python 3.3.1:"
	@~/.pythonbrew/pythons/Python-3.3.1/bin/python -m unittest discover

testpythonsall:
    # Test with all installed Python 2/3 releases.
	@echo "Testing with Python 2.6:"
	@~/.pythonbrew/pythons/Python-2.6/bin/python   -m          discover
	@echo "Testing with Python 2.7.2:"
	@~/.pythonbrew/pythons/Python-2.7.2/bin/python -m unittest discover
	@echo "Testing with Python 2.7.3:"
	@~/.pythonbrew/pythons/Python-2.7.3/bin/python -m unittest discover
	@echo "Testing with Python 2.7.4:"
	@~/.pythonbrew/pythons/Python-2.7.4/bin/python -m unittest discover
	@echo "Testing with Python 3.1:"
	@~/.pythonbrew/pythons/Python-3.1/bin/python   -m          discover
	@echo "Testing with Python 3.2:"
	@~/.pythonbrew/pythons/Python-3.2/bin/python   -m unittest discover
	@echo "Testing with Python 3.3.0:"
	@~/.pythonbrew/pythons/Python-3.3.0/bin/python -m unittest discover
	@echo "Testing with Python 3.3.1:"
	@~/.pythonbrew/pythons/Python-3.3.1/bin/python -m unittest discover

travistest:
	# Test runner for Travis-CI.
	@python -m discover

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

