#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys

from .basefile import PyHeaderFile

class Csv(PyHeaderFile):
    """
        class that read csv files with ; and , and #

        >>> test = Csv(name="test", header=["col1","col2","col3"])
        >>> test.write(*["test1","test2","test3"])
        >>> test.save('../')
        >>> test = Csv(name='../test.csv')
        >>> content = test.read()
        >>> sorted(next(content).items())
        [('col1', 'test1'), ('col2', 'test2'), ('col3', 'test3')]
        >>> test.name = 'test2'
        >>> from .excel import Xls, Xlsx
        >>> convert_xlsx = Xlsx()
        >>> convert_xlsx(test)
        >>> convert_xlsx.save()
        >>> convert_xls = Xls()
        >>> convert_xls(test)
        >>> convert_xls.save()

    """

    def __init__(self, name=None, header=list(), encode='utf-8', header_line=0,
                 delimiters=[",", ";", "#"], strip=False,
                 quotechar='"'):
        self.name = name
        self.header = header
        self.delimiters = list(delimiters)
        self.quotechar = quotechar
        self.encode = encode
        self.header_line = header_line
        self.strip = strip
        super(Csv, self).__init__()

    def read(self):
        # open file in mode write and read the line. Return the dict 'header = value'
        if isinstance(self.name, str) or isinstance(self.name, unicode):
            if not hasattr(self, '_file'):
                self._open()
            elif self._file.mode == 'w':
                self._file.close()
                self._open()
        else:
            self._open()
        for row in self.reader:
            if self.strip:
                row = [r.strip() for r in row]
            yield dict(zip(self.header, row))

    def write(self, *args, **kwargs):
        # write the value in the file
        if isinstance(self.name, str) or isinstance(self.name, unicode):
            if not hasattr(self, '_file'):
                self._file = open(self.name, 'a')
            elif self._file.mode == 'r':
                self._file.close()
                self._file = open(self.name, 'a')
        else:
            self._file = self.name
        writer = self.csv.DictWriter(self._file, delimiter=self.delimiters[0],
                                     fieldnames=self.header,
                                     quotechar=self.quotechar,
                                     quoting=self.csv.QUOTE_MINIMAL)
        if args:
            kwargs = dict(zip(self.header, args))
        writer.writerow(kwargs)

    def close(self):
        # close de file
        if not (isinstance(self.name, str) or isinstance(self.name, unicode)):
            return self.name.getvalue()

        if hasattr(self, '_file'):
            self._file.close()

    def save(self, path=None):
        # move and close de file
        if (isinstance(self.name, str) or isinstance(self.name, unicode)) and path:
            self.close()
            name = self.name
            basename = self.path.basename(name)
            self.rename(name, self.path.join(path, basename))
        elif path:
            name = 'default.csv'
            content = self.close()
            with open(self.path.join(path, name), 'w') as f:
                f.write(content)
        else:
           return self.close()

    def _get_dialect(self):
        # discover a dialect to csv file based on some delimiters
        try:
            for i in range(0, self.header_line):
                self._file.readline()
            self.dialect = self.csv.Sniffer().sniff(self._file.readline(),
                                                    delimiters=self.delimiters)
        except:
            self.dialect = self.delimiters[0]
        self._file.seek(0)

    def _open(self):
        # open the file and get header
        if isinstance(self.name, str) or isinstance(self.name, unicode):
            self._file = open(self.name, 'r')
        else:
            self._file = self.name
        self._file.seek(0)
        self._get_dialect()
        if (sys.version_info > (3, 0)):
            self.reader = self.csv.reader(self._file, self.dialect,
                                          doublequote=True)
        else:
            self.reader = self.csv.reader(self._file, self.dialect,
                                          encoding=self.encode, doublequote=True)
        for i in range(0, self.header_line):
            next(self.reader)
        self.header = next(self.reader)

    def _create(self):
        # create the file and write the header
        if isinstance(self.name, str) or isinstance(self.name, unicode):
            name = self.path.splitext(self.name)[0]
            self.name = "%s.csv" % name
            self._file = open(self.name, 'w')
        else:
            self._file = self.name

        if isinstance(self.name, str) or isinstance(self.name, unicode):
            self._file.seek(0)
            self.write(*self.header)
            self._file.close()
        else:
            self.write(*self.header)

    def _import(self):
        if (sys.version_info > (3, 0)):
            import csv
        else:
            import unicodecsv as csv
 
        import os

        self.rename = os.rename
        self.csv = csv
        self.path = os.path
