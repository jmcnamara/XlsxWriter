from distutils.core import setup

setup(
    name='XlsxWriter',
    version='0.5.1',
    author='John McNamara',
    author_email='jmcnamara@cpan.org',
    url='https://github.com/jmcnamara/XlsxWriter',
    packages=['xlsxwriter'],
    license='BSD',
    description='A Python module for creating Excel XLSX files.',
    long_description=open('README.rst').read(),
)
