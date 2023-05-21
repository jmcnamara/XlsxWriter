import sys
import os

extensions = ['sphinx.ext.intersphinx', 'sphinx.ext.extlinks']
source_suffix = '.rst'
master_doc = 'index'

project = u'XlsxWriter'
copyright = u'2013-2023, John McNamara'

version = '3.1.1'
release = version

exclude_patterns = []
pygments_style = 'sphinx'
intersphinx_mapping = {'python': ('https://docs.python.org/3', None)}

sys.path.append(os.path.abspath('_themes'))
html_theme_path = ['_themes']
html_theme = 'bootstrap'
html_theme_options = {'nosidebar': True}
html_title = "XlsxWriter Documentation"
html_static_path = ['_static']
html_show_sphinx = True
html_show_copyright = True
htmlhelp_basename = 'XlsxWriterdoc'
html_add_permalinks = ""

latex_elements = {
    'pointsize': '11pt',
    'preamble': '',
}
latex_documents = [
    ('index', 'XlsxWriter.tex',
     'Creating Excel files with Python and XlsxWriter',
     'John McNamara', 'manual'),
]

latex_logo = '_images/logo.png'
man_pages = [
    ('index', 'xlsxwriter',
     'XlsxWriter Documentation',
     ['John McNamara'], 1)
]

texinfo_documents = [
    ('index',
     'XlsxWriter',
     'XlsxWriter Documentation',
     'John McNamara',
     'XlsxWriter',
     'Creating Excel files with Python and XlsxWriter',
     'Miscellaneous'),
]

epub_title = 'XlsxWriter'
epub_author = 'John McNamara'
epub_publisher = 'John McNamara'
epub_copyright = '2013-2023, John McNamara'

linkcheck_ignore = [r'.*microsoft.com.*',
                    r'.*office.com.*',
                    r'.*www.paypal.com.*',
                    r'https://twitter.com/jmcnamara13']

extlinks = {'issue': ('https://github.com/jmcnamara/XlsxWriter/issues/%s', 'Issue #'),
            'feature': ('https://github.com/jmcnamara/XlsxWriter/issues/%s', 'Feature Request #'),
            'pull': ('https://github.com/jmcnamara/XlsxWriter/pull/%s', 'Pull Request #')}
