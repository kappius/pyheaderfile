#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
from setuptools import setup

setup(
  name = 'pyheaderfile',
  packages = ['pyheaderfile'],
  version = __import__('pyheaderfile').__version__,
  description = 'Enable handle of csv, xls and xlsx files getting '
                'column header',
  long_description = open(os.path.join(os.path.dirname(__file__), 'README.md'),
                          'r').read(),
  author = 'Diogo Munaro Vieira, Isvaldo Fernandes de Souza, '
           'Thiago Pereira Fernandes',
  author_email = 'diogo.mvieira@gmail.com, isvaldo.fernandes@gmail.com, '
                 'thiago.fernandes210@gmail.com',
  url = 'https://github.com/kappius/pyheaderfile',
  download_url = 'https://github.com/kappius/pyheaderfile/archive/v0.1.tar.gz',
  keywords = ['xls', 'excel', 'spreadsheet', 'workbook', 'xlsx', 'csv', 'txt'],
  license = 'Apache',
  classifiers = ['Programming Language :: Python',
                 'Programming Language :: Python :: 2',
                 'Programming Language :: Python :: 3',
                 'Operating System :: OS Independent',
                 'Topic :: Database',
                 'Topic :: Office/Business',
                 'Topic :: Software Development :: Libraries :: Python Modules',
                ],
)