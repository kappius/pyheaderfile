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

First of all you need to import module:

::

    from pyheaderfile import Csv, Xls, Xlsx, guess_type

Each of them will be explained below.


Class csv
---------

Read csv
^^^^^^^^

Default encode is utf8, but you can change it. Default strip is false, but classes can strip each value automatically:

::

    file = Csv(name=’file.csv’, encode='latin1', strip=True)
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

Save file
^^^^^^^^^

::

    file.save()

Class Xls
---------

Read xls
^^^^^^^^

You can strip automatically values from xls files too, but default value is False:

::

    file = Xls(name=’file.xls’, strip=True)
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

Save file
^^^^^^^^^

Finally you can save the file

::

    file.save()

Class Xlsx
----------

Read
^^^^

You can strip values from xlsx files too:

::

    file = Xlsx(name=’file.xlsx’, strip=True)
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

You can save the file to another path too

::

    file.save('/path/to/new/file/')

Alternativelly to save you can use close() that just use same path mandatorily.

::

    file.close()

Working with memory
-------------------

Writing
^^^^^^^

Objects can be stored in memory and then saved into disk or simple stay in memory:

::

    from StringIO import StringIO
    mem_obj = StringIO()
    xls = Xls(mem_obj, header=['first', 'second'])
    xls.write('1 guy', '2 guys')
    xls.save()  # or you can xls.save('/path/to/file/')

When you save file you retrieve StringIO contents or save its to disk specifying a directory. The content will be saved with name 'default.xls' in this case.


Reading
^^^^^^^

Same as writing you can read objects from memory. So, after you save your content you can read it again:

::

    from StringIO import StringIO
    mem_obj = StringIO()
    xls = Xls(mem_obj, header=['first', 'second'])
    xls.write('1 guy', '2 guys')
    xls.save()
    # here use new object
    new_xls = Xls(mem_obj)
    for row in new_xls:
        print row # should echo {'first': '1 guy', 'second': '2 guys'} then next rows

Tricks
------

Modifying extensions, name and header
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

You can change filename and header using this:

::

    q = Xls()
    x = Xlsx(name='filename.xlsx')
    x.name = 'ugly file name'
    x.header = ['col1', 'col2','col3']
    q(x)

BE CAREFUL! You can't change name using StringIO or others memory storage. You will get an error.

Guess file type
^^^^^^^^^^^^^^^

To guess what class you need to open just use:

::

    filename = 'test.xls'
    my_file = guess_type(filename)

If you are working with Csv or Xls, you can pass all possible kwargs and guess_type guess right kwargs:

::

    my_file = guess_type(filename, encode='latin1', strip=True)

Only if filename is a Csv file, then guess_type send encode kwarg to instance.

And for a SUPERCOMBO you can guess and convert everything!

::

    my_file = guess_type(filename, **kwargs)
    convert_to = Xls()
    my_file.name = 'beautiful_name'
    my_file.header = ['col1', 'col2','col3']
    convert_to(my_file) # now your file is a xls file ;)
    convert_to.save('/my/other/path/')
