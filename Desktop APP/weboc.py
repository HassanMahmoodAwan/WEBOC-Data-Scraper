import sys
from PyQt5.QtWidgets import (
     QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QLabel, QFileDialog, QLineEdit, QCheckBox, QSpacerItem, QSizePolicy
)
from PyQt5.QtGui import QFont, QIntValidator
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QTimer, QTime
import os
import shutil
from components import *
import threading

import asyncio
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup
import pandas as pd
import re
import openpyxl
from datetime import datetime




class MainWindow(QWidget):
    
    def __init__(self):
        self.numRecords = 100
        self.allRecords = False
        self.isFileUploaded = False
        self.exsistingPdfData = False
        self.fileName = ""
        
        self.totalCounter = 0
        self.progressCounter = 0
        
        super().__init__()
        self.initUI()
        
    
    
    
    def initUI(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
       
            
        # ****** Heading ******
        box1 = QWidget()  
        box1.setFixedHeight(75)  
        box1_layout = QVBoxLayout()  
        box1_layout.setAlignment(Qt.AlignCenter)  
        box1_layout.setContentsMargins(0,0,0,20)
        box1.setLayout(box1_layout)

        # Create and style the heading label
        label = QLabel("Weboc.gov.pk", self)
        label.setFont(QFont("Helvetica", 19, QFont.Bold))  
        label.setStyleSheet("color: black;")  
        label.setAlignment(Qt.AlignCenter) 
        
        box1_layout.addWidget(label)
        layout.addWidget(box1)
        # ===================== 
            
 
        
        
        # ********* Upload PDF Box *********
        box2 = QWidget()
        box2.setFixedHeight(90)
        box2.setFixedWidth(870)
        box2.setStyleSheet("""
            background-color: lightgray;
            border-radius: 15px;
        """)

        box2_layout = QHBoxLayout()
        box2.setLayout(box2_layout)

        self.file_label = QLabel("No file is uploaded", self)
        self.file_label.setFont(QFont("Helvetica", 12))
        self.file_label.setStyleSheet("color: black; padding-left: 12px;")

        upload_button = QPushButton("Upload File", self)
        upload_button.setFont(QFont("Helvetica", 11))
        upload_button.setStyleSheet("""
        QPushButton {
            background-color: darkblue;
            color: white;
            border-radius: 10px;
            padding: 12px 24px 9px 24px; 
            margin-right: 4px; 
        }
        QPushButton:hover {
            background-color: blue;  
        }
        """)
        upload_button.clicked.connect(self.upload_pdf)

        box2_layout.addWidget(self.file_label)
        box2_layout.addStretch()
        box2_layout.addWidget(upload_button)
        
        layout.addWidget(box2, alignment=Qt.AlignHCenter)
        # ======================================    
        
        
        
        
        # ************ Number of HsCode, ALL, Existing HsCode ****************
        box3 = QWidget()
        box3.setFixedHeight(50)
        box3.setFixedWidth(870)           
               
        box3_layout = QHBoxLayout() 
        box3.setLayout(box3_layout)

        self.number_input = QLineEdit(self)
        self.number_input.setValidator(QIntValidator())
        self.number_input.setPlaceholderText("Enter a number")
        

        self.all_checkbox = QCheckBox("ALL", self)

        self.hs_codes_checkbox = QCheckBox("Use existing HsCodes", self)

        box3_layout.addWidget(self.number_input)
        box3_layout.addWidget(self.all_checkbox)
        
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        box3_layout.addItem(spacer)

        box3_layout.addWidget(self.hs_codes_checkbox)

        layout.addWidget(box3, alignment=Qt.AlignHCenter)
        # ==================================
        
        
        
        
        
        
        # ********** RUN Btn Box **********
        box4 = QWidget()
        box4.setFixedHeight(70) 
        box4.setFixedWidth(870)
        box4_layout = QVBoxLayout()
        box4_layout.setContentsMargins(0, 0, 0, 0)
        box4.setLayout(box4_layout)

        self.button_box4 = QPushButton("Submit and Run!", self)
        self.button_box4.clicked.connect(self.runApp)
        self.button_box4.setFont(QFont("Helvetica", 13))
        self.button_box4.setStyleSheet("""
            QPushButton {
                background-color: darkblue;
                color: white;
                border-radius: 8px;
                padding: 12px 20px 9px 20px;
                width: 100%; 
            }
            QPushButton:hover {
                background-color: blue;
            }
        """)
        box4_layout.addWidget(self.button_box4)  

        layout.addWidget(box4, alignment=Qt.AlignHCenter)
        # =============================================
        
        
    
              
        # *********** Status Labels ***********
        box5 = QWidget()
        box5.setFixedWidth(870)
        box5.setFixedHeight(150)

        box5_layout = QVBoxLayout()                
        box5.setLayout(box5_layout)

        self.pdf_extraction_label = QLabel("Pdf Extraction:    Not Started", self)
        self.pdf_extraction_label.setFont(QFont("Helvetica", 10))
        self.pdf_extraction_label.setStyleSheet("color: black;")

        self.hs_code_label = QLabel("HsCode:             ", self)
        self.hs_code_label.setFont(QFont("Helvetica", 10))
        self.hs_code_label.setStyleSheet("color: black;")

        self.status_label = QLabel("Status:                No Record", self)
        self.status_label.setFont(QFont("Helvetica", 10))
        self.status_label.setStyleSheet("color: black;")

        self.counter_label = QLabel(f"Counter:             {self.progressCounter} / {self.totalCounter}", self)
        self.counter_label.setFont(QFont("Helvetica", 10))
        self.counter_label.setStyleSheet("color: black;")
        
        self.timer_label = QLabel("Timer:      00:00:00", self)
        self.status_label.setFont(QFont("Helvetica", 10))
        self.timer_label.setStyleSheet("color: black; font-weight: bold;")
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_timer)


        box5_layout.addWidget(self.pdf_extraction_label)
        box5_layout.addWidget(self.hs_code_label)
        box5_layout.addWidget(self.status_label)
        box5_layout.addWidget(self.counter_label)
        box5_layout.addWidget(self.timer_label)


        layout.addWidget(box5, alignment=Qt.AlignHCenter)
        # ==============================================
        
        
        
        # ******** Success Error Box ************
        self.box6 = QWidget()
        self.box6.setFixedHeight(55)
        self.box6.setFixedWidth(870)
        self.box6.setStyleSheet("""
            
            border-radius: 10px;
        """)
        
        box6_layout = QVBoxLayout()
        self.box6.setLayout(box6_layout)

        self.box6_label = QLabel("", self)
        self.box6_label.setFont(QFont("Helvetica", 13))
        self.box6_label.setStyleSheet("color: black; padding-left: 15px;")

        box6_layout.addWidget(self.box6_label)

        layout.addWidget(self.box6, alignment=Qt.AlignHCenter)
        # ===========================================
        
        
        
        
        
        self.setLayout(layout)
        self.setWindowTitle("Weboc - Data Scraper App")
        self.setGeometry(550, 300, 1000, 630)
        self.show()
        
 


    def update_timer(self):
        self.time = self.time.addSecs(1)
        self.timer_label.setText(f"Timer:      {self.time.toString('hh:mm:ss')}")
    def stop_timer(self):
        self.timer.stop()


    # ********** Functions Started ************ 
    
    # UPload PDF (Main Thread)
    def upload_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open PDF File", "", "PDF Files (*.pdf)")
        
        self.box6_label.setText("")
        self.box6.setStyleSheet("")
        
        if file_path:
            save_folder = "uploaded_pdfs"
            Excel_folder = "Excel-Files"
            if not os.path.exists(save_folder):
                os.makedirs(save_folder)
                
            if not os.path.exists(Excel_folder):
                os.makedirs(Excel_folder)
            
            self.fileName = os.path.basename(file_path)
            
            save_path = os.path.join(save_folder, self.fileName)
            shutil.copy(file_path, save_path)
            print(f"File saved to: {save_path}")
            
            self.file_label.setText(f"{self.fileName}")
            self.isFileUploaded = True
            
     
            
    def runApp(self):
        
        self.status_label.setText(f"Status:                No Record")
        self.hs_code_label.setText(f"HsCode:             ")
        self.progressCounter = 0
        self.totalCounter = 0
        self.counter_label.setText(f"Counter:             {self.progressCounter} / {self.totalCounter}")
        
        
        if self.all_checkbox.isChecked():
            self.allRecords = self.all_checkbox.isChecked()
            self.box6_label.setText("")
            self.box6.setStyleSheet("")
        elif self.number_input.text() != "":
            self.allRecords = self.all_checkbox.isChecked()
            self.numRecords = int(self.number_input.text())
            self.box6_label.setText("")
            self.box6.setStyleSheet("")
        else:
            self.box6_label.setText("Provide Number of Records")
            self.box6.setStyleSheet("""
            background-color: red;
            border-radius: 10px;""")
            self.box6_label.setStyleSheet("color: white; padding-left: 15px;")
            return
         
        
           
        if self.hs_codes_checkbox.isChecked():
            self.exsistingPdfData = True
            self.box6_label.setText("")
            self.box6.setStyleSheet("")
        else:
            if self.isFileUploaded:
                self.box6_label.setText("")
                self.box6.setStyleSheet("")
            else:
                self.box6_label.setText("No File is Uploaded or no Record Exist")
                self.box6.setStyleSheet("""
                background-color: red;
                border-radius: 10px;""")
                self.box6_label.setStyleSheet("color: white; padding-left: 15px;")
                return

        
        
        self.time = QTime(0, 0, 0)
        self.timer_label.setText("Timer:      00:00:00")
        self.timer.start(1000)
        
            
        
        #  *************************** WEB Scraping CODE ****************************
        async def scraper(hsCodeList:list[str], isAllPages = True, onlyOneRow = False, maxPagesAllowed:int = 5):
            # CLearing Existing Data
            file_name = 'weboc_data.xlsx'
            outputExcelPath = f"./Excel-Files/{file_name}"
            outputExcelPath = os.path.abspath(outputExcelPath)
            if os.path.exists(outputExcelPath):
                os.remove(outputExcelPath)  
                print(f"{file_name} has been deleted.")
            else:
                print(f"{file_name} does not exist.")



            web_url = "https://www.weboc.gov.pk/(S(p4qc02boyxm1t1bc2mjszqta))/DownloadValuationData.aspx"
            progressCounter = 0

            async with async_playwright() as playwright:
                browser = await playwright.chromium.launch(headless=False, args=["--no-sandbox"])  
                context = await browser.new_context()
                page = await context.new_page()
                await page.goto(web_url)

                for hsCode in hsCodeList:
                    
                    
                    if type(hsCode) != str or len(hsCode) != 9:
                        self.progressCounter += 1
                        self.counter_label.setText(f"Counter:             {self.progressCounter} / {self.totalCounter}")
                        continue
                    print(hsCode) 
                    self.hs_code_label.setText(f"HsCode:             {hsCode}") 
                    
                    self.progressCounter += 1
                    self.counter_label.setText(f"Counter:             {self.progressCounter} / {self.totalCounter}")  

                    await page.fill('#txtHSCode', hsCode)
                    await page.click('#btnSearch')
                    await page.wait_for_timeout(1000)



                    try:
                        element = await page.wait_for_selector('#dgList', timeout=1000, state="attached")
                        print("Records found")
                        self.status_label.setText("Status:                Record Found")
                        
                    except:
                        checkStr = await page.inner_text("#lblMessage", timeout=500)
                        print(checkStr)
                        self.status_label.setText(f"Status:                {checkStr}")
                        continue
                    
                    
                    if isAllPages == False:
                        pageContent = await page.content()
                        data = extract_data(pageContent, hsCode, onlyOneRow)   
                        format_data(data)

                    else:

                        num_Pages = await page.inner_text("#ctrlPageRender_lblPageDetails")
                        num_Pages = int(num_Pages.split(" ")[-1])

                        counter = 1
                        # while(counter <= num_Pages and counter <= maxPagesAllowed):
                        while(counter <= num_Pages):

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
                self.stop_timer()
                if self.progressCounter == self.totalCounter:
                        self.box6_label.setText(f"Successfully, Scraped All the Data.")
                        self.box6.setStyleSheet("""
                            background-color: lightGreen;
                            border-radius: 10px;""")
                        self.box6_label.setStyleSheet("color: black; padding-left: 15px;")
                
        

        
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
            df = pd.DataFrame.from_dict(data, orient='index', columns=['id', "HS Code",'Description', 'Unit value',  "Goods Name" , 'Country'])
            df = df.drop(columns=['id'])
            
            current_DateTime = datetime.now()
            dateStamp = current_DateTime.strftime("%d-%m-%Y") 
            df["Date"] = dateStamp
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


                
        
        def run(hsCodeList:list[str], isAllPages = True, onlyOneRow = False, maxPagesAllowed:int = 5):
            isAllPages = True
            maxPagesAllowed = 5
            asyncio.run(scraper(hsCodeList, isAllPages=isAllPages, onlyOneRow=False, maxPagesAllowed=maxPagesAllowed))
            # Database Connection here
        #  **************************************************************************
        
        
               
        
        
        # App Runner
        def runner(allRecord = False, numRecords = 100, existingHsCode = False, isFileUploaded = False, fileName:str = ""):
            
            self.pdf_extraction_label.setText("Pdf Extraction:    In Progress")
            
            if fileName != "":
                if allRecord == True:
                    result = pdf_extractor(pdf_name=fileName, newExtraction=not (existingHsCode), allRecords=allRecord) 
                else: 
                    result = pdf_extractor(pdf_name=fileName, count=numRecords, newExtraction=not (existingHsCode), allRecords=False)             
            else: 
                if allRecord == True:
                    result = pdf_extractor(newExtraction=not (existingHsCode), allRecords=allRecord) 
                else:
                    result = pdf_extractor(count=numRecords, newExtraction= not(existingHsCode), allRecords=False)
                
                
                
            if type(result) == list:
                self.pdf_extraction_label.setText("Pdf Extraction:    Completed")
                self.totalCounter = len(result)
                self.counter_label.setText(f"Counter:             {self.progressCounter} / {self.totalCounter}")
                self.box6_label.setText("")
                self.box6.setStyleSheet("")
                # isAllPages = True
                self.thread = threading.Thread(target=run, args=([result]))
                self.thread.start()
            else:
                self.pdf_extraction_label.setText("Pdf Extraction:    Error")
                self.box6_label.setText(f"{result}")
                self.box6.setStyleSheet("""
                background-color: red;
                border-radius: 10px;""")
                self.box6_label.setStyleSheet("color: white; padding-left: 15px;")
                return
                
      
      
        thread = threading.Thread(target=runner, args=(self.allRecords, self.numRecords, self.hs_codes_checkbox.isChecked(), self.isFileUploaded, self.fileName))
        thread.start()



def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())  # Use exec_() for PyQt5


if __name__ == "__main__":
    main()
