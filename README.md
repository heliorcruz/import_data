# import_data
Import data from .xlsx to MongoDB
The example.xlsl can be used as template for the data to be imported

- settings configuration config.py   
CONN_STRING = 'mongodb://admin:admin@127.0.0.1'
- name of db and collection to store   
DATABASE = 'teste'  
COLLECTION = 'teste'  
- file name to import  
FILENAME = 'example.xlsx'  
  
ALFA = "ABCDEFGHIJKLMNOPQRSTUVXZ"  
- field collumns   
FIELDS = [ "DATA",
           "NOME", 
           "ENDERECO", 
           "NUM", 
           "COMPLEMENTO", 
           "DATA_NASC", 
           "TELEFONE", 
           "CELULAR",  
           "CEP", 
           "CIDADE", 
           "UF", 
           "EMAIL", 
           "DATA"]
- categories  
CATEG  = ["GRUPO_1", "GRUPO_2"]  
- number of collumns  
COLLUMS = 12  
- number of rows in .xlslx  
ROWS    = 1000  
- position of header   
HEADER  = 1  
- save to db with upsert  
UPSERT  = False  

