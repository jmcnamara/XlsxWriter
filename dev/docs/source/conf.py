import sys
import os

extensions = ['sphinx.ext.intersphinx', 'sphinx.ext.extlinks']
source_suffix = '.rst'
master_doc = 'index'

project = u'XlsxWriter'
copyright = u'2013-2023, John McNamara'

version = '3.1.9'
release = version

exclude_patterns = []
intersphinx_mapping = {'python': ('https://docs.python.org/3', None)}


html_title = "XlsxWriter"
html_show_sphinx = True
html_show_copyright = True
#html_static_path = ['_static']
#html_css_files = ["custom.css"]

html_theme = 'pydata_sphinx_theme'


html_theme_options = {
    "navbar_align": "left",
    "header_links_before_dropdown": 1,
    "secondary_sidebar_items": ["page-toc"],
    "navbar_end": [],
    "navbar_center": ["navbar-nav"],

    "pygment_light_style": "vs",
    "pygment_dark_style": "monokai"
}


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

linkcheck_ignore = ["https://www.paypal.com"]

extlinks = {'issue': ('https://github.com/jmcnamara/XlsxWriter/issues/%s', 'Issue %s'),
            'feature': ('https://github.com/jmcnamara/XlsxWriter/issues/%s', 'Feature Request %s'),
            'pull': ('https://github.com/jmcnamara/XlsxWriter/pull/%s', 'Pull Request %s')}
