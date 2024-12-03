import pandas as pd
import os

# # # Define file paths
# Path = os.getcwd() + "/Documents/Excel-Files/temps.xlsx"
# output_file_path = os.getcwd() + "/Documents/Output_temp.xlsx"

# # Check if the file exists
# if os.path.exists(Path):
#     # Load the DataFrame
#     df = pd.read_excel(Path, index_col=None, header=None, dtype={0: str})
#     print("Initial DataFrame:")
    
#     # Print the initial shape
#     print("Initial shape:", df.shape)

#     # Remove duplicates, ignoring the 'Date' column
#     # df = df.loc[~df.drop(columns=['Date']).duplicated(keep='first')]
#     df.drop_duplicates(subset=None, keep='first', inplace=True)
#     df = df.drop(index=0).reset_index(drop=True)
#     print(df.head(6))

    

#     # Print the final shape
#     print("Shape after removing duplicates:", df.shape)

#     # Save the updated DataFrame to an Excel file
#     df.to_excel(output_file_path, index=False, engine='openpyxl')
#     print(f"File saved to {output_file_path}")



# Load two dataframes and merge the second one into first and save in Output.xlsx.
# Path_one = os.getcwd() + "/Documents/Output.xlsx"
# Path_two = os.getcwd() + "/Documents/Output_temp.xlsx"
output_file_path = os.getcwd() + "/Output.xlsx"


# df_one  = pd.read_excel(Path_one, index_col=None, dtype={'HS Code': str})
# df_two =  pd.read_excel(Path_two, index_col=None, header=None, dtype={0: str})
# df_two = df_two.drop(index=0).reset_index(drop=True)
# df_two.columns = df_one.columns

# print(df_one.shape)
# print(df_two.shape)
# merged_df = pd.concat([df_one, df_two], ignore_index=True)
# merged_df.to_excel(output_file_path, index=False, engine='openpyxl')
# print("Saved")

# print(merged_df.head(5))
# print(df_one.head(5))
# print(df_two.head(5))



df_one  = pd.read_excel(output_file_path, index_col=None, dtype={'HS Code': str})
df = df_one.drop(columns='Date')
df.drop_duplicates(subset=None, keep='first', inplace=True)

df["Date"] = "03-12-2024"



print(df.shape)
print(df.head())


df.to_excel(output_file_path, index=False, engine='openpyxl')
print("Successful")