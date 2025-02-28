# STAGE_PHONES.py
# STAGE_PHONES.py

import pandas as pd
import os

# File path
file_path = r"C:\Users\us85360\OneDrive - Grant Thornton LLP\0 - All Work\[c] 15 - Unitil\0_SM Work\BangorDataSAPSampleExtract.xlsx"

# Read the Excel file and load the specific sheet
df = pd.read_excel(file_path, sheet_name='ZCAMPAIGN', engine='openpyxl')

# Initialize df_new using relevant columns
df_new = pd.DataFrame()

# Assign CUSTOMERID and PHONENUMBER with proper formatting
df_new['CUSTOMERID'] = df.iloc[:, 1].fillna('').astype(str).str.slice(0, 15)
df_new['PHONENUMBER'] = df.iloc[:, 8].fillna('').astype(str).str.slice(0, 10)

# Add PHONETYPE, ensuring it has 2 digits (leading zeros if needed)
df_new['PHONETYPE'] = '01'  # Default value
df_new['PHONETYPE'] = df_new['PHONETYPE'].astype(str).str.zfill(2)

# Add additional columns with default empty values
additional_columns = ['PHONEEXT', 'CONTACT', 'TITLE', 'PRIORITY', 'UPDATEDATE']
for col in additional_columns:
    df_new[col] = " "

# Add trailer row
trailer_row = pd.DataFrame([['TRAILER'] + [','] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Define output file path
output_csv = os.path.join(os.path.dirname(file_path), 'STAGE_PHONE.csv')

# Save to CSV
df_new.to_csv(output_csv, index=False)

# Confirmation message
print(f"CSV file saved at {output_csv}")
