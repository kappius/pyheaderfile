#!/usr/bin/env python
# -*- coding: utf-8 -*-


class PyHeaderFile(object):
    # father class of all filetypes

    def __init__(self, *args, **kwargs):
        self._import()
        if self.name:
            if self.header:
                self._create()
            self._open()

    def __call__(self, instance, **kwargs):

        # convert any File object to any File and save it

        self.__init__(instance.name, instance.header, **kwargs)

        for line in instance.read():
            if isinstance(line[instance.header[0]], tuple):
                new_line = dict()
                for l in line:
                    new_line[l] = line[l][0]
                self.write(**new_line)
            else:
                self.write(**line)
        self.save()

    def read(self):

        # read each line of file. Should be an interator

        raise NotImplementedError

    def write(self, *args, **kwargs):
        """
        write to file by args or kwargs. Should be a list
        with elements or a dict for writes with unordered
        lines
        """
        raise NotImplementedError

    def __exit__(self):
        self.save()

    def save(self, path=None):

        # save and close file in other path

        return NotImplemented

    # can use method close or save
    def close(self):

        # save and close file

        return NotImplemented


    # getter and setter for filename. You can change filename to convert
    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, name):
        self._name = name

    # getter and setter for file header
    @property
    def header(self):
        return self._header

    @header.setter
    def header(self, header):
        self._header = header

    def _open(self):
        """
        open file and get header of file. Some oder needed initialization
        things can be done here
        """
        raise NotImplementedError

    def _create(self):

        # create new file with right extension

        raise NotImplementedError

    def _import(self):

        # import needed libs. This prevent conflicts

        raise NotImplementedError


class PyHeaderSheet(PyHeaderFile):
    # class that use similar functions for sheets
    def __init__(self):
        self._row = 0
        super(PyHeaderSheet, self).__init__()
        if self.header:
            if isinstance(self.header[0], tuple):
                self.header = [h[0] for h in self.header]
        if not self.sheet_name and self.name:
            self._first_sheet()
        self._open_sheet()

    # define and get sheet name into a spreadsheet file
    @property
    def sheet_name(self):
        return self._sheet_name

    @sheet_name.setter
    def sheet_name(self, sheet_name):
        self._sheet_name = sheet_name

    def _first_sheet(self):
        # get first sheet
        try:
            self.sheet_name = self.sheet_names[0]
        except IndexError:
            raise Exception('There are no sheets')

    @property
    def sheet_names(self):
        # returns a list with the sheet file
        return self._sheet_names

    @sheet_names.setter
    def sheet_names(self, sheet_names):
        # returns a list with the sheet file
        self._sheet_names = sheet_names

    def read(self):
        # read the file line
        for x in range(1, self.nrows):
            row = dict()
            for y in range(0, self.ncols):
                row.update(self.read_cell(x, y))
            yield row

    # pass x, y, value and style for function write_cell
    def write(self, *args, **kwargs):
        """
        :param args: tuple(value, style), tuple(value, style)
        :param kwargs: header=tuple(value, style), header=tuple(value, style)
        :param args: value, value
        :param kwargs: header=value, header=value
        """

        if args:
            kwargs = dict(zip(self.header, args))
        for header in kwargs:
            cell = kwargs[header]
            if not isinstance(cell, tuple):
                cell = (cell,)
            self.write_cell(self._row, self.header.index(header), *cell)
        self._row += 1
