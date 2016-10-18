#!/usr/bin/env python
# -*- coding: utf-8 -*-

from test_basefile import TestBaseFile

from pyheaderfile.pyheaderfile.excel import Xls, Xlsx


class TestXls(TestBaseFile):

    def setUp(self):
        self.name = 'test.xls'
        self.klass = Xls


class TestXlsx(TestBaseFile):

    def setUp(self):
        self.name = 'test.xlsx'
        self.klass = Xlsx
