import sqlite3
from time import strptime
from openpyxl import load_workbook
import pandas as pd
import re



#file_path = input('Enter path to BMS - ')
#file_path = r'C:\Users\1\Documents\BMS\MIELE_2015_goods.xls'
#file_path = r'C:\Users\1\Documents\BMS\3M_2016_goods.xlsx'
file_path = r'C:\Users\1\Documents\BMS\Tsukorargoprom_2016_SugarProduction.xlsx'
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
nace_codes = ','.join(re.compile('\s([0-9]+)\s').findall(str(df3.iloc[7,1])))
search_date = df3.iloc[4,2]


########## WORKING WITH DATABASE ###############
connect = sqlite3.connect('bms_db.sqlite')
cur = connect.cursor()

cur.executescript('''

        CREATE TABLE IF NOT EXISTS Search(
            id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
            search_name TEXT UNIQUE,
            client TEXT,
            year TEXT,
            search_type TEXT,
            database TEXT,
            nace_codes TEXT,
            search_date TEXT
            );

        CREATE TABLE IF NOT EXISTS Company(
            id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
            printscreen TEXT UNIQUE,
            company TEXT,
            edrpou TEXT,
            nace_code TEXT,
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

cur.execute('''INSERT OR IGNORE INTO Search (search_name, client, year, search_type, database, nace_codes, search_date) VALUES (?,?,?,?,?,?,?)''',
            (search_name, client,year, search_type, database, nace_codes, search_date))
cur.execute('''SELECT id FROM Search WHERE search_name = ?''', (search_name,))
search_id = cur.fetchone()[0]


search_date = search_date.replace('201','1')
search_date = strptime(search_date, '%d/%m/%y')
#company_df = df1.iloc[:,['Название принтскрина','Компания','NACE Rev. 2.','ЕДРПОУ/ИИН','Адрес в интернете','Перечень деятельности','Описание деятельности','Причина отклонения','Источник информации','Источник информации 2']]
company_df = df1.iloc[:,[1,2,4,5,10,11,18,19,20,21]] ####!!!!to replace with links to normal column names!!!!
#print(company_df.columns)
for item in range(len(company_df.index)):
    row = list(company_df.iloc[item])
    printscreen = row[0]
    company = row[1]
    nace_code = row[2]
    edrpou = row[3]
    input_website = row[4]
    description_by_bvd = row[5]
    description = row[6]
    rejection_reason = row[7]
    found_website = row[8]
    found_website2 = row[9]

    print(item,company,type(search_date))

    
    try:
        cur.execute('''SELECT search_date FROM Company WHERE printscreen = ?''', (printscreen,))
        search_date_previous = cur.fetchone()[0]
        search_date_previous = search_date_previous.replace('201', '1')
        search_date_previous = strptime(search_date_previous, '%d/%m/%y')
        if search_date_previous < search_date:
            cur.execute('''INSERT OR REPLACE INTO Company
             (printscreen, company, edrpou, nace_code, input_website,description_by_bvd, description, 
             rejection_reason,found_website, found_website2, search_id, search_date) 
             VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''', (printscreen,
             company,
                                                   edrpou,
                                                   nace_code,
                                                   input_website,
                                                   description_by_bvd,
                                                   description,
                                                   rejection_reason,
                                                   found_website,
                                                   found_website2,
                                                   search_id,
                                                   str(search_date)))
    except:

        cur.execute('''INSERT OR IGNORE INTO Company
                     (printscreen, company, edrpou, nace_code, input_website,description_by_bvd, description, 
                     rejection_reason,found_website, found_website2, search_id, search_date) 
                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''', (printscreen,
                                                           company,
                                                           edrpou,
                                                           nace_code,
                                                           input_website,
                                                           description_by_bvd,
                                                           description,
                                                           rejection_reason,
                                                           found_website,
                                                           found_website2,
                                                           search_id,
                                                           str(search_date)))

connect.commit()
