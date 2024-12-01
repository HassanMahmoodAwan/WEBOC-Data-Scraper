from components import *
import pandas as pd
# import pyodbc
import os


HsCodes = pdf_extractor(pdf_name = "PAKISTANCUSTOMSTARIFF-2023-24.pdf", count=10, newExtraction=False, allHsCode=True)
print(HsCodes)
if type(HsCodes) == list:
    # HsCodes.insert(5, "1511.9030")
    # print(HsCodes)
    with open(os.getcwd() + "/Documents/hsCode.txt", "r") as f:
            code = f.read()
            print(code)
    if bool(code):
        run(HsCodes, isAllPages=True, maxPagesAllowed=5, isContinue=True, existingHsCode=code)
    # run(HsCodes, isAllPages=True, maxPagesAllowed=5, isContinue=True, existingHsCode=code)
    
    # ****** Dumping Data into DB. *******
    # file_name = 'weboc_data.xlsx'
    # sourceData_filePath = f"./Documents/Excel-Files/{file_name}"
    # sourceData_filePath = os.path.abspath(sourceData_filePath)
    
    # df = pd.read_excel(sourceData_filePath, engine='openpyxl')

    
    # server = 'localhost'  
    # database = 'SalesDB'
    # table_name = 'YourTableName'  
    # username = 'your_username' 
    # password = 'your_password' 
    # driver = '{ODBC Driver 17 for SQL Server}'
    
    
    # try:
    #     conn = pyodbc.connect(
    #         f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}Trusted_Connection=yes;')
    #     print("Connected to SQL Server successfully!")
        
    #     df.to_sql(table_name, con=conn, if_exists='append', index=False, method='multi')
    #     print(f"Data appended successfully to {table_name}!")

    # except Exception as e:
    #     print("An error occurred while connecting or inserting data: ", e)
    
    # finally:
    #     conn.close()
    #     print("Connection closed.")
    
    
else:
    print(HsCodes)