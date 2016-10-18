#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import inspect

from .excel import Xls, Xlsx
from .headercsv import Csv


def guess_type(filename, **kwargs):
    """ Utility function to call classes based on filename extension.
    Just usefull if you are reading the file and don't know file extension.
    You can pass kwargs and these args are passed to class only if they are
    used in class.
    """

    extension = os.path.splitext(filename)[1]
    case = {'.xls': Xls,
            '.xlsx': Xlsx,
            '.csv': Csv}
    if extension and case.get(extension.lower()):
        low_extension = extension.lower()
        new_kwargs = dict()
        class_name = case.get(low_extension)
        class_kwargs = inspect.getargspec(class_name.__init__).args[1:]
        for kwarg in kwargs:
            if kwarg in class_kwargs:
                new_kwargs[kwarg] = kwargs[kwarg]
        return case.get(low_extension)(filename, **new_kwargs)
    else:
        raise Exception('No extension found')
