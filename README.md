# PyHeaderFile

##Objetivo
O PyHeaderFile tem o objetivo de facilitar a vida com arquivos com extensões txt, csv, xls e xlsx. 
Esse projeto facilita a leitura de arquivos através da header, a conversão de extensões,
a retirada dos espaços de cada valor das linhas, a modificação dos stylos das celulas dos arquivos excel, a ler uma celula especifica no arquivo excel e a ler o arquivo setando a linha em que a header se encontra. 


* Como usa
    * Read

```
Csv(self, name=’file.csv’, header_line=0, delimiters=[",", ";", "#"], strip=False, quotechar='"')
```