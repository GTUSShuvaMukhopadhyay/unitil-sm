#STAGE_MAIL_ADDR.py

import pandas as pd
import os
import re

# File path
file_path = r"C:\Users\us85360\OneDrive - Grant Thornton LLP\0 - All Work\[c] 15 - Unitil\0_SM Work\BangorDataSAPSampleExtract.xlsx"

# Read the Excel file
df = pd.read_excel(file_path, sheet_name='MAILING_ADDR1', engine='openpyxl')

# Initialize new DataFrame
df_new = pd.DataFrame()

# ACCOUNTNUMBER (Column 8)
df_new['ACCOUNTNUMBER'] = df.iloc[:, 1].astype(str).str.slice(0, 15)

# ADDRESSSEQ (Auto-increment sequence)
df_new['ADDRESSSEQ'] = range(1, len(df) + 1)

# MAILINGNAME
df_new['MAILINGNAME'] = df.apply(
    lambda row: str(row.iloc[4]) + str(row.iloc[5]) if row.iloc[16] == 1 else row.iloc[2], 
    axis=1
).astype(str).str.slice(0, 35)

# INCAREOF (Column 6)
df_new['INCAREOF'] = df.iloc[:, 6].astype(str).str.slice(0, 35)

# ADDRESS1 (Column 7, fallback to Column 9 if empty)
df_new['ADDRESS1'] = df.iloc[:, 7].apply(
    lambda x: str(x)[:35] if pd.notna(x) and x != '' else str(df.iloc[:, 9])[:35]
)

# ADDRESS2 (Column 8)
df_new['ADDRESS2'] = df.iloc[:, 8].astype(str).str.slice(0, 35)

# CITY (Column 10)
df_new['CITY'] = df.iloc[:, 10].astype(str).str.slice(0, 24)

# STATE (Column 11)
df_new['STATE'] = df.iloc[:, 11].astype(str).str.slice(0, 2)

# COUNTRY (Hardcoded)
df_new['COUNTRY'] = "US"

# POSTALCODE - Preserve ZIP+4 format (e.g., 04416-1864)
df_new['POSTALCODE'] = df.apply(
    lambda row: str(row.iloc[12])[:15] if pd.notna(row.iloc[9]) and row.iloc[9] != '' and re.match(r'^[a-zA-Z0-9\-]*$', str(row.iloc[9])) 
    else str(row.iloc[13])[:15],
    axis=1
)

# Ensure 'POSTALCODE' is a string, strip spaces
df_new['POSTALCODE'] = df_new['POSTALCODE'].fillna('').astype(str).str.strip()

# Add leading zero if exactly 4 characters
df_new['POSTALCODE'] = df_new['POSTALCODE'].apply(lambda x: '0' + x if len(x) == 4 else x)

# Validate POSTALCODE (Only allow alphanumeric and hyphen)
df_new['POSTALCODE'] = df_new['POSTALCODE'].apply(lambda x: x if re.match(r'^[a-zA-Z0-9\-]*$', x) else '')

# UPDATEDATE (Empty for now)
df_new['UPDATEDATE'] = ""

# Function to wrap values in double quotes except ADDRESSSEQ (numeric) and UPDATEDATE (date)
def quote_wrap(x):
    if isinstance(x, (int, float)):  # Keep numbers unchanged
        return x  
    return f'"{x}"' if pd.notna(x) else '""'  # Wrap in quotes

# Apply quote wrapping to all columns except ADDRESSSEQ and UPDATEDATE
for col in df_new.columns:
    if col not in ['ADDRESSSEQ', 'UPDATEDATE']:
        df_new[col] = df_new[col].astype(str).apply(quote_wrap)

# Add trailer row
trailer_row = pd.DataFrame([['TRAILER'] + [','] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Define output file path
output_csv = os.path.join(os.path.dirname(file_path), 'STAGE_MAIL_ADDR.csv')

# Save to CSV
df_new.to_csv(output_csv, index=False)

# Confirmation message
print(f"CSV file saved at {output_csv}")
