PyHeaderFile
************

The PyHeaderFile helps the work with files that have extensions csv, xls and xlsx.

This project aims **reading files over the header (column names)**. With this module we can handle **Csv, Xls and Xlsx files using same interface**. Thus, we can convert extensions, strip values in lines, change cell style of Excel files, read a specific Excel file, read an specific cell and read just some headers.

Install
=======

::

    pip install pyheaderfile

How to use
==========

Class csv
---------

Read csv
^^^^^^^^

::

    file = Csv(name=’file.csv’)
    for row in file.read():
        print row  


Set Header
^^^^^^^^^^

::

    file.header = ['col1', 'col2','col3']


Create csv
^^^^^^^^^^

::

    file = Csv(name='filename.csv', header=['col1','col2','col3'])


Write list csv
^^^^^^^^^^^^^^

::

    file.write(['col1','col2','col3'])


Write dict csv
^^^^^^^^^^^^^^

::

    file.write(dict(header=value))

Class Xls
---------

Read xls
^^^^^^^^

::

    file = Xls(name=’file.xls’)
    for row in file.read():
        print row  


Set Header
^^^^^^^^^^

::

    file.header = ['col1', 'col2','col3']


Create xls
^^^^^^^^^^

::

    file = Xls(name='filename.xls', header=['col1','col2','col3'])


Write list
^^^^^^^^^^

::

    file.write(['col1','col2','col3'])


Write dict
^^^^^^^^^^

::

    file.write(dict(header=value))


Class Xlsx
----------

Read
^^^^

::

    file = Xlsx(name=’file.xlsx’)
    for row in file.read():
        print row  


Set Header
^^^^^^^^^^

::

    file.header = ['col1', 'col2','col3']


Create file
^^^^^^^^^^^

::

    file = Xlsx(name='filename.xlsx', header=['col1','col2','col3'])


Write list
^^^^^^^^^^

::

    file.write(['col_val1','col_val2','col_val3'])


Write dict
^^^^^^^^^^

::

    file.write(dict(header=value))


Save file
^^^^^^^^^

::

    file.save()

Tricks
------

Modifying extensions, name and header
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

::

    q = Xls()
    x = Xlsx(name='filename.xlsx')
    x.name = 'ugly file name'
    x.header = ['col1', 'col2','col3']
    q(x)


Guess file type
^^^^^^^^^^^^^^^

To guess what class you need to open just use:

::

    filename = 'test.xls'
    my_file = guess_type(filename)

And for a SUPERCOMBO you can guess and convert everything!

::

    my_file = guess_type(filename)
    convert_to = Xls()
    my_file.name = 'beautiful_name'
    my_file.header = ['col1', 'col2','col3']
    convert_to(my_file) # now yout file is a xls file ;)
