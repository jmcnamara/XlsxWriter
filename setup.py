import sys
from warnings import warn

from setuptools import setup

if sys.version_info < (3, 8):
    warn("The minimum Python version supported by XlsxWriter is 3.8")
    sys.exit()

setup(
    name="xlsxwriter",
    version="3.2.8",
    author="John McNamara",
    author_email="jmcnamara@cpan.org",
    url="https://github.com/jmcnamara/XlsxWriter",
    packages=["xlsxwriter"],
    scripts=["examples/vba_extract.py"],
    license="BSD-2-Clause",
    description="A Python module for creating Excel XLSX files.",
    long_description=open("README.rst", encoding="utf-8").read(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "License :: OSI Approved :: BSD License",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3 :: Only",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
    ],
    python_requires=">=3.8",
)
