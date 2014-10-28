# PyHeaderFile

O PyHeaderFile tem o objetivo de facilitar a sua vida, com arquivos que tenham extensões csv, xls e xlsx. 
Esse projeto visa facilitar a leitura de arquivos através da header. Com as classes Csv, Xls e Xlsx é possivel 
a conversão de extensões, retirada dos espaços de cada valor das linhas, a modificação dos stylos das celulas
dos arquivos excel, a ler uma celula especifica no arquivo excel e a ler o arquivo setando a linha em que a header se encontra. 

## Install

## Como usa
* Class csv

Read csv
```
file = Csv(name=’file.csv’)
for row in file.read():
    print row  
```

Modificando a Header
```
file.header = list('coluna1', 'coluna2','coluna3')
```

Criando um csv
```
file = Csv(name='filename.csv', header=['coluna1','coluna2','coluna3'])
```

Write list csv
```
file.write(list('coluna1','coluna2','coluna3'))
```

write dict csv
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

Modificando a Header
```
file.header = list('coluna1', 'coluna2','coluna3')
```

Criando um xls
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

Modificando a Header
```
file.header = list('coluna1', 'coluna2','coluna3')
```

Criando um arquivo
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

* Modificando extenções
```
q = Xls()
x = Xlsx(name='filename.xlsx')
x.name = 'file'
q(x)
```
