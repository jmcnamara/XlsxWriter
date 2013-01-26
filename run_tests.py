#
# Simple test runner to allow automation.
#
# Same as: $ python -m unittest discover
#
import unittest


def load_tests(loader, tests, pattern):
    all_tests = unittest.TestSuite()
    for suites in unittest.defaultTestLoader.discover('.', pattern='test*.py'):
        for tests in suites:
            all_tests.addTests(tests)
    return all_tests

if __name__ == '__main__':
    unittest.main()
