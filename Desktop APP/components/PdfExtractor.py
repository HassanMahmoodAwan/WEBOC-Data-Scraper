import pdfplumber
import pandas as pd
import os

def pdf_extractor(pdf_name:str = "PAKISTANCUSTOMSTARIFF-2023-24.pdf", count:int= 20, newExtraction = False, allRecords = False) -> list | str:

    
    outputExcelPath = "./Excel-Files/pdfData.xlsx"
    outputExcelPath = os.path.abspath(outputExcelPath)
    
    if newExtraction == False:
        try:
            df = pd.read_excel(outputExcelPath)
            if allRecords == True:
                return df["HsCode"].tolist()
            return df.head(count)["HsCode"].tolist()
        except Exception as e:
            return f"Existing HsCode Data not Found !"
        

    
    combined_df = pd.DataFrame()
    pdf_path:str = f"./uploaded_pdfs/{pdf_name}"
    pdf_path =  os.path.abspath(pdf_path)
    print(pdf_path)
    
    try:
        print("****** PDF Data Extraction Started ******")
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:                
                    df = pd.DataFrame(table[1:])
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                
            print("Data Extracted Sucessfully!") 
    
    except FileNotFoundError:
         return f"PDF file not Exsist: {e}"
    except Exception as e:
        return f"PDF not extracted, Error Occured: {e}"
    else:
        combined_df.columns = ["HsCode", "Description", "CD (%)"] if int(combined_df.shape[1]) == 3 else combined_df
        combined_df.to_excel(outputExcelPath, index=False)
        print("No of Rows: ", combined_df.shape[0])
    
    try:
        if allRecords == True:
                return combined_df["HsCode"].tolist()
        return combined_df.head(count)["HsCode"].tolist()
    except Exception as e:
        print("Not converted into List: ", e)
        return []
    


