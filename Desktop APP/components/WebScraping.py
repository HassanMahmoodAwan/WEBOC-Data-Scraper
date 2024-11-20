import asyncio
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup
import pandas as pd
import re
import openpyxl
import os


async def scraper(hsCodeList:list[str], isAllPages = False, onlyOneRow = False, maxPagesAllowed:int = 5):
    
    web_url = "https://www.weboc.gov.pk/(S(p4qc02boyxm1t1bc2mjszqta))/DownloadValuationData.aspx"
    progressCounter = 0
    
    
                       
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=False)  
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto(web_url)
            
        for hsCode in hsCodeList:
            if type(hsCode) != str or len(hsCode) != 9:
                continue
            print(hsCode)    

            await page.fill('#txtHSCode', hsCode)
            await page.click('#btnSearch')
            await page.wait_for_timeout(1000)

            
            
            try:
                element = await page.wait_for_selector('#dgList', timeout=1000, state="attached")
                print("Records found")
            except:
                checkStr = await page.inner_text("#lblMessage", timeout=500)
                print(checkStr)
                continue
            
             
            if isAllPages == False:
                pageContent = await page.content()
                data = extract_data(pageContent, hsCode, onlyOneRow)   
                format_data(data)

            else:

                num_Pages = await page.inner_text("#ctrlPageRender_lblPageDetails")
                num_Pages = int(num_Pages.split(" ")[-1])
        
                counter = 1
                while(counter <= num_Pages and counter <= maxPagesAllowed):

                    await page.fill('#ctrlPageRender_txtGoToPage', str(counter))
                    await page.click('#ctrlPageRender_btnGoTo')
                    await asyncio.sleep(1)
                    await page.wait_for_selector('#dgList tbody tr', timeout=10000)

                    pageContent = await page.content()

                    # await asyncio.sleep(0.5)
                    data = extract_data(pageContent, hsCode)   
                    format_data(data)  

                    counter += 1
                    num_Pages = await page.inner_text("#ctrlPageRender_lblPageDetails")
                    num_Pages = int(num_Pages.split(" ")[-1])
                    print("PageCount:", num_Pages)
                    print("counter value:", counter)
            
            
        await browser.close()
        
        
        
        
        
        
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
         
        counter += 1 
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
            
            print(goodName_match.group()) if goodName_match else print(None)
            
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
    save_data(df)


# ****** Saving Records in Excel File *******
def save_data(df):
    file_name = 'weboc_data.xlsx'
    outputExcelPath = f"./Excel-Files/{file_name}"
    outputExcelPath = os.path.abspath(outputExcelPath)

    try:
        reader = pd.read_excel(outputExcelPath)
        writer = pd.ExcelWriter(outputExcelPath, engine='openpyxl', mode='a', if_sheet_exists='overlay') 
        df.to_excel(writer, index=False, header=False, startrow=len(reader) + 1)
        writer.close()

    except FileNotFoundError:
        df.to_excel(outputExcelPath, index=False)






# ******** Function runner *******
def run(hsCodeList:list[str], isAllPages = False, onlyOneRow = False, maxPagesAllowed:int = 5):
    asyncio.run(scraper(hsCodeList, isAllPages, onlyOneRow, maxPagesAllowed))
