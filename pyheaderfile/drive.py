#!/usr/bin/env python
# -*- coding: utf-8 -*-

from .basefile import PyHeaderSheet

class GSheet(PyHeaderSheet):
    """
    Class that read google spreadsheet files

    """

    def __init__(self, email, password, name=None, header=list(),
                 sheet_name=None, strip=False):
        """
        :param email: email
        :param password: password
        :param name: file name
        :param header: list with the header ['text1','text2','text3']
        :param sheet_name: sheet name
        :param strip: true to take spaces of values
        :return:
        """
        self.name = name
        self.email = email
        self.password = password
        self.header = header
        self.strip = strip
        self.sheet_name = sheet_name
        super(GSheet, self).__init__()


    def read_cell(self, x, y):
        """
        Reads the cell at position x+1 and y+1; return value
        :param x: line index
        :param y: coll index
        :return: {header: value}
        """
        if isinstance(self.header[y], tuple):
            header = self.header[y][0]
        else:
            header = self.header[y]
        x += 1
        y += 1
        if self.strip:
            self._sheet.cell(x, y).value = self._sheet.cell(x, y).value.strip()
        else:
            return {header: self._sheet.cell(x, y).value}

    def write_cell(self, x, y, value):
        """
        Writing value in the cell of x+1 and y+1 position
        :param x: line index
        :param y: coll index
        :param value: value to be written
        :return:
        """
        x += 1
        y += 1
        self._sheet.update_cell(x, y, value)

    def save(self, path=None):
        """

        :param path:
        :return:
        """
        if path:
            raise NotImplementedError
        return

    def _open(self):
        """
        Open the file; get sheets
        :return:
        """
        if not hasattr(self, '_file'):
            self._file = self.gc.open(self.name)
            self.sheet_names = self._file.worksheets()

    def _open_sheet(self):
        """
        Read the sheet, get value the header, get number columns and rows
        :return:
        """
        if self.sheet_name and not self.header:
            self._sheet = self._file.worksheet(self.sheet_name.title)
            self.ncols = self._sheet.col_count
            self.nrows = self._sheet.row_count
            for i in range(1, self.ncols+1):
                self.header = self.header + [self._sheet.cell(1, i).value]

    def _create(self):
        """

        :return:
        """
        raise NotImplementedError

    def _import(self):
        """
        Makes imports
        :return:
        """
        import os.path
        import gspread
        self.path = os.path
        self.gspread = gspread
        self._login()

    def _login(self):
        """
        Login with your Google account
        :return:
        """
        # TODO(dmvieira) login changed to oauth2
        self.gc = self.gspread.login(self.email, self.password)
