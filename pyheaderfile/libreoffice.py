#!/usr/bin/env python
# -*- coding: utf-8 -*-

from .basefile import PyHeaderSheet

# TODO(dmvieira) not implemented
class Ods(PyHeaderSheet):
    """
    class that read ods files.
    Need reimplementing with new module. Peharps theses links:
    http://www.marco83.com/work/173/read-an-ods-file-with-python-and-odfpy/
    http://opendocumentfellowship.com/projects/odfpy
    """

    def __init__(self, name=None, header=list(), sheet_name=None):
        self.name = name
        self.header = header
        self.sheet_name = sheet_name
        super(Ods, self).__init__()

    def read_cell(self, x, y):
        raise NotImplementedError

    def write_cell(self, x, y, value, style=None):
        raise NotImplementedError

    def save(self, path=None):
        raise NotImplementedError

    def close(self, path=None):
        raise NotImplementedError

    def get_sheets(self):
        # self.sheets = [s.name for s in self._file.sheets]
        # return self.sheets
        return NotImplementedError

    def _open(self):
        # self._file = self.ezodf.opendoc(self.name)
        raise NotImplementedError

    def _import(self):
        # import ezodf
        # self.ezodf = ezodf
        raise NotImplementedError
