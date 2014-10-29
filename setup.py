try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

setup(
    name='XlsxWriter',
    version='0.6.1',
    author='John McNamara',
    author_email='jmcnamara@cpan.org',
    url='https://github.com/jmcnamara/XlsxWriter',
    packages=['xlsxwriter'],
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
    ],
)
