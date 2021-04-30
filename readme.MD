
# 1. Setup
Tested with python 3.9.

1. Install dependencies with `pip install -r requirements.txt`  


## 2. Usage

With a space seperated list:
```
.\analyse.py transactions.csv --retailers 'Lidl' 'Aldi'
```

Or by passing a file (one retailer on every line):

```
.\analyse.py transactions.csv --retailers-file retailers.csv
```


Or by passing a file list (one file on every line):  
note: assumes file name relative to execution location.

```
.\analyse.py file_list.csv --multi --retailers-file retailers.csv
```