#!/usr/bin/env python
# -*- coding: utf-8 -*-

__all__ = ['Csv', 'Xls', 'Xlsx', 'guess_type']

from pyheaderfile.headercsv import *
from pyheaderfile.excel import *
from pyheaderfile.libreoffice import *
from pyheaderfile.drive import *
from pyheaderfile.utils import *

VERSION = (0, 3, 0)
__version__ = ".".join(map(str, VERSION))
