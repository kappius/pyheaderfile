#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys

from .test_basefile import TestBaseFile


class TestCsv(TestBaseFile):

    def test_should_write(self):
        if (sys.version_info > (3, 0)):
            expected = 'col1,col2,col3\ntest1,test2,test3\n'
        else:
            expected = 'col1,col2,col3\r\ntest1,test2,test3\r\n'
        test = self.klass(name=self.name, header=["col1", "col2", "col3"])
        test.write(*["test1", "test2", "test3"])
        test.close()
        with open(self.name) as csv_file:
            self.assertEqual(expected, csv_file.read())

    def test_should_read(self):
        content = 'col1,col2,col3\r\ntest1,test2,test3\r\n'
        with open(self.name, 'w') as csv_file:
            csv_file.write(content)

        test = self.klass(name=self.name)
        content = test.read()
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        self.assertDictEqual(expected, next(content))
        test.close()
