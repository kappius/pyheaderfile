#!/usr/bin/env python
# -*- coding: utf-8 -*-


def guess_type(filename,**kwargs):
    """ Utility function to call classes based on filename extension.
    Just usefull if you are reading the file and don't know file extension.
    You can pass kwargs and these args are passed to class only if they are
    used in class.
    """
    import os

    extension = os.path.splitext(filename)[1]
    case = {'.xls': Xls,
            '.xlsx': Xlsx,
            '.csv': Csv}
    if extension and case.get(extension.lower()):
        low_extension = extension.lower()
        new_kwargs = dict()
        class_name = case.get(low_extension)
        class_kwargs = class_name.__init__.func_code.co_names
        for kwarg in kwargs:
            if kwarg in class_kwargs:
                new_kwargs[kwarg] = kwargs[kwarg]
        return case.get(low_extension)(filename, **new_kwargs)
    else:
        raise Exception('No extension found')


def is_str_or_unicode(value):
    """
    Verifies if sting or unicode.
    :param value: value to be verified
    :return: True or None
    """
    if isinstance(value, str) or isinstance(value, unicode):
        return True

################################################################################
# run tests
################################################################################

if __name__ == '__main__':
    import doctest
    doctest.testmod()
