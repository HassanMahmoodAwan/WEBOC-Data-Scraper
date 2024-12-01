import asyncio
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup
import pandas as pd
import re
import openpyxl
import os
from datetime import datetime
import time

# ===== Global Variable ======
dataSavingCounter = 0
singleHsCodeData = {}

async def scraper(hsCodeList:list[str], isAllPages = False, onlyOneRow = False, maxPagesAllowed:int = 5, isContinue=False, existingHsCode = None):
    global dataSavingCounter
    global singleHsCodeData
    
    # Check Existing HsCode exist to Continue.
    if isContinue:  
        try:
            start_index =  hsCodeList.index(existingHsCode) 
            dataSavingCounter = len(pd.read_excel(os.getcwd() + "/Documents/Excel-Files/weboc_data.xlsx"))
        except:
            print("HsCode donot exist in List")
            return
    else:
        start_index = 0
        dataSavingCounter = 0


    web_url = "https://www.weboc.gov.pk/(S(p4qc02boyxm1t1bc2mjszqta))/DownloadValuationData.aspx"
    
    # new Extraction, Delete existing Data.
    if not isContinue:
        outputExcelPath = f"./Documents/Excel-Files/weboc_data.xlsx"
        outputExcelPath = os.path.abspath(outputExcelPath)
        if os.path.exists(outputExcelPath):
            os.remove(outputExcelPath)  
            print("weboc_data.xlsx has been deleted.")
        else:
            print("weboc_data.xlsx does not exist.")
     
    # try:                      
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=False)  
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto(web_url)
        for index in range(start_index, len(hsCodeList)):
            hsCodeData = {}
            if type(hsCodeList[index]) != str or len(hsCodeList[index]) != 9:
                continue
            print(hsCodeList[index])   
            with open(os.getcwd() + "/Documents/hsCode.txt", "w") as f:
                f.write(hsCodeList[index])
                f.close()
            await page.fill('#txtHSCode', hsCodeList[index])
            await page.click('#btnSearch')
            await page.wait_for_timeout(1500)
            # chech record found, if not, continue.
            try:
                element = await page.wait_for_selector('#dgList', timeout=1500, state="attached")
                print("Records found")
            except:
                checkStr = await page.inner_text("#lblMessage", timeout=500)
                print(checkStr)
                continue
            
            
            if isAllPages == False:
                pageContent = await page.content()
                data = extract_data(pageContent, hsCodeList[index], onlyOneRow)   
                format_data(data)
            else:
                num_Pages = await page.inner_text("#ctrlPageRender_lblPageDetails")
                num_Pages = int(num_Pages.split(" ")[-1])
                counter = 1
                # while(counter <= num_Pages and counter <= maxPagesAllowed):
                while(counter <= num_Pages):
                    print("Counter")
                    await page.fill('#ctrlPageRender_txtGoToPage', str(counter))
                    await asyncio.sleep(0.2)
                    await page.click('#ctrlPageRender_btnGoTo')
                    print("Clicked")
                    await asyncio.sleep(1)
                    await page.wait_for_selector('#dgList tbody tr', timeout=100000, state="attached")
                    pageContent = await page.content()
                    await asyncio.sleep(1)
                    data = extract_data(pageContent, hsCodeList[index], onlyOneRow) 
                    hsCodeData = singleHsCodeData =  {**hsCodeData, **data}
                      
                    counter += 1
                    num_Pages = await page.inner_text("#ctrlPageRender_lblPageDetails")
                    num_Pages = int(num_Pages.split(" ")[-1])
                    print("PageCount:", num_Pages)
                    print("counter value:", counter)
                
                format_data(hsCodeData)
                hsCodeData = {}
                singleHsCodeData = {}
        await browser.close()
    # except Exception as e:
    #     print("Any Issue is Occured, might be Internet Issue.")
    #     print("Exception: ", e)
        
        # format_data(singleHsCodeData)
        
        
        
        
        
        
        
        
        
#  ****** Extract data using BeautifulSoup *******
def extract_data(htmlContent: str, hsCode: str, isOnlyOneRow = False) -> dict:
    soup = BeautifulSoup(htmlContent, "html.parser")
    
    table = soup.find('table', {'id': 'dgList'})
    tbody = table.find('tbody')
    data = {}
    counter = 0
    for tr in tbody.find_all('tr'):
        if "HeaderStyle" in tr.get("class", []):
            continue  
         
        row_data = [td.get_text(strip=True) for td in tr.find_all('td')]
        
        for value in data.values():
            if value[2] == row_data[1] and value[5] == row_data[2]:
                break
        else:
            
            match = re.search(r'(\d+\.\d+\s+\w+\s*)[^\w]*$', row_data[1]) 
            row_data.insert(2, match.group(0)) if match else row_data.insert(2, None)
            
            
            #  ******* Extracting the Goods Name *******
            if row_data[1][0] in ['H', 'h']:
                goodName_match = re.search(r'^(.*?)(?=HsCode)', row_data[1])
            else:
                goodName_match = re.match(r'^[^H]*(?=HsCode)', row_data[1])
            
            bracket_match = re.match(r'^[^(]*', goodName_match.group()) if goodName_match else None
            if bracket_match:
                dot_match = re.match(r'^[^.]+', bracket_match.group())
            else:
                dot_match = re.match(r'^[^.]+', goodName_match.group()) if goodName_match else None
            
            if dot_match:
                goodName_match = dot_match
            else:    
                goodName_match = goodName_match if goodName_match else None   
            # ==========================
                       
            
            row_data.insert(3, goodName_match.group(0)) if goodName_match else row_data.insert(3, None)
            row_data.insert(1, hsCode)
            data[row_data[0]] = row_data

            if isOnlyOneRow == True:
                break; 

    return data



# ****** Creating DataFrame and Saving ******
def format_data(data):
    df = pd.DataFrame.from_dict(data, orient='index', columns=['id', "HS Code",'Description', 'Unit value', "Goods Name" , 'Country'])
    df = df.drop(columns=['id'])

    current_DateTime = datetime.now()
    dateStamp = current_DateTime.strftime("%d-%m-%Y") 
    df["Date"] = dateStamp
    save_data(df)


# ****** Saving Records in Excel File *******
def save_data(df):
    global dataSavingCounter
    file_name = 'weboc_data.xlsx'
    outputExcelPath = f"./Documents/Excel-Files/{file_name}"
    outputExcelPath = os.path.abspath(outputExcelPath)

    
    print("current Counter: ", dataSavingCounter)
    
    try:
        # reader = pd.read_excel(outputExcelPath)
        writer = pd.ExcelWriter(outputExcelPath, engine='openpyxl', mode='a', if_sheet_exists='overlay') 
        df.to_excel(writer, index=False, header=False, startrow=dataSavingCounter + 1)
        print("writed into excel file")
        writer.close()
        
        dataSavingCounter += len(df)
    except FileNotFoundError:
        df.to_excel(outputExcelPath, index=False)



async def retry_exception_Runner(hsCodeList:list[str], isAllPages = False, onlyOneRow = False, maxPagesAllowed:int = 5, isContinue=False, existingHsCode = None):
    retry_counter= 0
    retry_attempts = 100
    retry_interval = 15                 # 15 Secs break. 
    
    
    while (retry_counter <= retry_attempts):
        try:
            await scraper(hsCodeList, isAllPages, onlyOneRow, maxPagesAllowed, isContinue, existingHsCode)
            print("Successfully, retreive all the data.")
            with open(os.getcwd() + "/Documents/hsCode.txt", "w") as f:
                f.write("")
                f.close()
            break
        except Exception as e:
            print("An Exception Occurred: ", e)
            print("Excecution Failed")
            
            format_data(singleHsCodeData)
        
            with open(os.getcwd() + "/Documents/hsCode.txt", "r") as f:
                existingHsCode = f.read()
            
            retry_counter += 1
            
            if retry_counter >= retry_attempts:
                print("Max Retry attempts Reached. ")
                break
            
            print("Retrying Program Execution in {}.".format(retry_interval))
                
            await asyncio.sleep(retry_interval)




# ******** Function runner *******
def run(hsCodeList:list[str], isAllPages = False, onlyOneRow = False, maxPagesAllowed:int = 5, isContinue=False, existingHsCode = None):
    
    asyncio.run(retry_exception_Runner(hsCodeList, isAllPages, onlyOneRow, maxPagesAllowed, isContinue, existingHsCode))
    
            
            
            
            
