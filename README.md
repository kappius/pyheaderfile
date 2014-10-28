# PyHeaderFile

The PyHeaderFile aims to facilitate the work with files that have extensions csv , xls and xlsx.
This project aims to facilitate reading files over the header. With this classes Csv , Xls and Xlsx is possible
converting extensions, removal of areas of lines of each value , the modification of the cell stylos
the excel file , read a specific excel file and read the file by setting the line in the header cell is .

## Install
```pip install pyheaderfile```

## How to use
* Class csv

Read csv
```
file = Csv(name=’file.csv’)
for row in file.read():
    print row  
```

Set Header
```
file.header = list('coluna1', 'coluna2','coluna3')
```

Create csv
```
file = Csv(name='filename.csv', header=['coluna1','coluna2','coluna3'])
```

Write list csv
```
file.write(list('coluna1','coluna2','coluna3'))
```

Write dict csv
```
file.write(dict(header=value))
```
* Class Xls

Read xls
```
file = Xls(name=’file.xls’)
for row in file.read():
    print row  
```

Set Header
```
file.header = list('coluna1', 'coluna2','coluna3')
```

Create xls
```
file = Xls(name='filename.xls', header=['coluna1','coluna2','coluna3'])
```

Write list
```
file.write(list('coluna1','coluna2','coluna3'))
```

Write dict
```
file.write(dict(header=value))
```

* Class Xlsx

Read
```
file = Xlsx(name=’file.xlsx’)
for row in file.read():
    print row  
```

Set Header
```
file.header = list('coluna1', 'coluna2','coluna3')
```

Create file
```
file = Xlsx(name='filename.xlsx', header=['coluna1','coluna2','coluna3'])
```

Write list
```
file.write(list('coluna1','coluna2','coluna3'))
```

Write dict
```
file.write(dict(header=value))
```

Save file
```
file.save()
```

* Modifying extensions
```
q = Xls()
x = Xlsx(name='filename.xlsx')
x.name = 'file'
q(x)
```
