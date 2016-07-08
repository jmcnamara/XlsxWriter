import sys
import subprocess
from warnings import warn

try:
    from setuptools import setup, Command
except ImportError:
    from distutils.core import setup, Command

if sys.version_info < (2, 5, 0):
    warn("The minimum Python version supported by XlsxWriter is 2.5.")
    exit()


class PyTest(Command):

    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        errno = subprocess.call(['python',  '-m', 'unittest', 'discover'])
        raise SystemExit(errno)

setup(
    name='XlsxWriter',
    version='0.9.3',
    author='John McNamara',
    author_email='jmcnamara@cpan.org',
    url='https://github.com/jmcnamara/XlsxWriter',
    packages=['xlsxwriter'],
    scripts=['examples/vba_extract.py'],
    cmdclass={'test': PyTest},
    license='BSD',
    description='A Python module for creating Excel XLSX files.',
    long_description=open('README.rst').read(),
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.5',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.1',
        'Programming Language :: Python :: 3.2',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
    ],
)
