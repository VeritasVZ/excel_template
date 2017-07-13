import sqlite3
import datetime
from openpyxl import load_workbook
import pandas as pd
import re

connect = sqlite3.connect('bms_db.sqlite')
cur = connect.cursor()

cur.executescript('''
DROP TABLE IF EXISTS Search;
DROP TABLE IF EXISTS Company;

CREATE TABLE Search(
    id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
    search_name TEXT UNIQUE,
    client TEXT,
    year TEXT,
    search_type TEXT,
    database TEXT,
    nace_codes TEXT,
    search_date TEXT
    );

CREATE TABLE Company(
        id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
        printscreen TEXT UNIQUE,
        company TEXT,
        edrpou TEXT,
        nace_code INTEGER,
        input_website TEXT,
        description_by_bvd TEXT,
        description TEXT,
        rejection_reason TEXT,
        found_website TEXT,
        found_website2 TEXT,
        search_id INTEGER,
        search_date TEXT
        );
''')

#file_path = input('Enter path to BMS - ')
file_path = r'C:\Users\1\Documents\BMS\MIELE_2015_goods.xls'
file = pd.ExcelFile(file_path)
df1 = file.parse('Проверка деятельности')
#df2 = file.parse('Результаты') >>> not relevant untill we add  table with financial data
df3 = file.parse('Стратегия поиска')

search_properties = re.compile('[a-zA-Z0-9]+').findall(file_path)
client = search_properties[-4]
year = search_properties[-3]
search_type = search_properties[-2]
search_name = client+'_'+year+'_'+search_type
database = list(df3.columns)[2]
#print(df3.iat[7,1])
nace_codes = ','.join(re.compile('\s([0-9]+)\s').findall(str(df3.iloc[7,1])))
search_date = df3.iloc[4,2]
#print(df3)
#print(client)
#print(year)
#print(search_type)
#print(database)
#print(nace_codes)
#print(search_date)

cur.execute('''INSERT OR IGNORE INTO Search (search_name) VALUES (?)''', (search_name,))
cur.execute('''SELECT id FROM Search WHERE search_name = ?''', (search_name,))
search_id = cur.fetchone()[0]

cur.execute('''INSERT OR REPLACE INTO Search (search_name, client, year, search_type, database, nace_codes, search_date) VALUES (?,?,?,?,?,?,?)''',
            (search_name, client,year, search_type, database, nace_codes, search_date))

#company_df = df1.iloc[:,['Название принтскрина','Компания','NACE Rev. 2.','ЕДРПОУ/ИИН','Адрес в интернете','Перечень деятельности','Описание деятельности','Причина отклонения','Источник информации','Источник информации 2']]
company_df = df1.iloc[:,[2,3,5,6,11,12,19,20,21,22]]
for item in range(len(company_df.index)):
    row = list(company_df.iloc[item])
    printscreen = row[0]
    company = row[1]
    edrpou = row[2]
    nace_code = row[3]
    input_website = row[4]
    description_by_bvd = row[5]
    description = row[6]
    rejection_reason = row[7]
    found_website = row[8]
    found_website2 = row[9]

    try:
        cur.execute('''SELECT search_date FROM Company WHERE printscreen = ?''', (printscreen,))
        search_date_previous = cur.fetchone()[0]
    except:
        search_date_previous = search_date
    if search_date_previous is not None:
        #search_date_previous.replace('/','-')
        datetime.datetime.strptime(search_date_previous, '%d/%m/%y').date()
        print(search_date_previous, type(search_date_previous))
        #print(search_date,type(search_date))
        if search_date_previous > search_date:
            print('hooray')
            continue
    #cur.execute('''INSERT OR ''')

connect.commit()
