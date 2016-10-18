#!/usr/bin/env python
# -*- coding: utf-8 -*-


def is_str_or_unicode(value):
    """
    Verifies if sting or unicode.
    :param value: value to be verified
    :return: True or None
    """
    if isinstance(value, str) or isinstance(value, unicode):
        return True
