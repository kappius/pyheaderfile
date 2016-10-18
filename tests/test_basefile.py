import os

from unittest import TestCase

from pyheaderfile.pyheaderfile import Xls, Xlsx
from pyheaderfile.pyheaderfile import Csv
from pyheaderfile.pyheaderfile import guess_type


class TestBaseFile(TestCase):

    def setUp(self):
        self.name = 'test.csv'
        self.klass = Csv

    def test_should_write_and_read(self):
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        test = self.klass(name=self.name, header=["col1", "col2", "col3"])
        test.write(*["test1", "test2", "test3"])
        test.close()
        test = self.klass(name=self.name)
        content = test.read()
        self.assertDictEqual(expected, next(content))
        test.close()

    def test_should_write_and_read_guessing_type(self):
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        test = self.klass(name=self.name, header=["col1", "col2", "col3"])
        test.write(*["test1", "test2", "test3"])
        test.close()
        test = guess_type(self.name)
        content = test.read()
        self.assertDictEqual(expected, next(content))
        test.close()

    def test_should_write_with_no_extension_and_read(self):
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        test = self.klass(name=self.name.split('.')[0], header=["col1", "col2", "col3"])
        test.write(*["test1", "test2", "test3"])
        test.close()
        test = self.klass(name=self.name)
        content = test.read()
        self.assertDictEqual(expected, next(content))
        test.close()

    def test_should_write_and_read_from_another_path(self):
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        test = self.klass(name=self.name, header=["col1", "col2", "col3"])
        test.write(*["test1", "test2", "test3"])
        test.save('../')
        self.name = '../{}'.format(self.name)
        test = self.klass(name=self.name)
        content = test.read()
        self.assertDictEqual(expected, next(content))
        test.close()

    def test_should_write_and_read_using_dict(self):
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        test = self.klass(name=self.name, header=["col1", "col2", "col3"])
        test.write(**{'col2': 'test2', 'col1': 'test1', 'col3': 'test3'})
        test.close()
        test = self.klass(name=self.name)
        content = test.read()
        self.assertDictEqual(expected, next(content))
        test.close()

    def test_should_write_and_convert_to_csv(self):
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        test = self.klass(name=self.name, header=["col1", "col2", "col3"])
        test.write(**{'col2': 'test2', 'col1': 'test1', 'col3': 'test3'})
        test.close()
        test = self.klass(name=self.name)
        convert = Csv()
        convert(test)
        convert.close()
        test = Csv(name=self.name.split('.')[0] + '.csv')
        content = test.read()
        self.assertDictEqual(expected, next(content))
        test.close()
        os.remove(self.name.split('.')[0] + '.csv')

    def test_should_write_and_convert_to_xls(self):
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        test = self.klass(name=self.name, header=["col1", "col2", "col3"])
        test.write(**{'col2': 'test2', 'col1': 'test1', 'col3': 'test3'})
        test.close()
        test = self.klass(name=self.name)
        convert = Xls()
        convert(test)
        convert.close()
        test = Xls(name=self.name.split('.')[0] + '.xls')
        content = test.read()
        self.assertDictEqual(expected, next(content))
        test.close()
        os.remove(self.name.split('.')[0] + '.xls')

    def test_should_write_and_convert_to_xlsx(self):
        expected = {'col1': 'test1', 'col2': 'test2', 'col3': 'test3'}
        test = self.klass(name=self.name, header=["col1", "col2", "col3"])
        test.write(**{'col2': 'test2', 'col1': 'test1', 'col3': 'test3'})
        test.close()
        test = self.klass(name=self.name)
        convert = Xlsx()
        convert(test)
        convert.close()
        test = Xlsx(name=self.name.split('.')[0] + '.xlsx')
        content = test.read()
        self.assertDictEqual(expected, next(content))
        test.close()
        os.remove(self.name.split('.')[0] + '.xlsx')

    def tearDown(self):
        try:
            os.remove(self.name)
        except:
            pass
