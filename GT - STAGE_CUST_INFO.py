# STAGE_CUST_INFO.py

import pandas as pd
import os

# File path
file_path = r"C:\Users\us85360\OneDrive - Grant Thornton LLP\0 - All Work\[c] 15 - Unitil\0_SM Work\BangorDataSAPSampleExtract.xlsx"

# Read the Excel file and load the specific sheet
df = pd.read_excel(file_path, sheet_name='MAILING_ADDR1', engine='openpyxl')

# Initialize df_new as an empty DataFrame
df_new = pd.DataFrame()

# Column 1: CUSTOMERID from Column B (index 1)
df_new['CUSTOMERID'] = df.iloc[:, 1].fillna('').astype(str).str.slice(0, 15)

# Column 2: FULLNAME based on conditions
df_new['FULLNAME'] = df.apply(
    lambda row: f"{str(row.iloc[4])}{str(row.iloc[5])}" if row.iloc[16] == 1 else str(row.iloc[2]),
    axis=1
).str.slice(0, 50)

# Column 3: FIRSTNAME from Column E (index 4)
df_new['FIRSTNAME'] = df.iloc[:, 4].fillna('').astype(str).str.slice(0, 25)

# Empty MIDDLENAME
df_new['MIDDLENAME'] = " "

# Column 4: LASTNAME from Column F (index 5)
df_new['LASTNAME'] = df.iloc[:, 5].fillna('').astype(str).str.slice(0, 25)

# Placeholder columns
df_new['NAMETITLE'] = " "
df_new['NAMESUFFIX'] = " "
df_new['DBA'] = " "

# Column 6: CUSTTYPE from Column Q (index 16)
df_new['CUSTTYPE'] = df.iloc[:, 16].fillna('').astype(str).str.slice(0, 1)

# Column 7: ACTIVECODE
df_new['ACTIVECODE'] = "0"

# Additional columns
additional_columns = [
    'MOTHERMAIDENNAME', 'EMPLOYERNAME', 'EMPLOYERPHONE', 'EMPLOYERPHONEEXT',
    'OTHERIDTYPE1', 'OTHERIDVALUE1', 'OTHERIDTYPE2', 'OTHERIDVALUE2',
    'OTHERIDTYPE3', 'OTHERIDVALUE3', 'UPDATEDATE'
]
for col in additional_columns:
    df_new[col] = " "

# Function to wrap values in double quotes except CUSTTYPE (numeric) and ACTIVECODE (numeric)
def quote_wrap(x):
    if isinstance(x, (int, float)):  # Keep numbers unchanged
        return x  
    return f'"{x}"' if pd.notna(x) else '""'  # Wrap in quotes

# Apply quote wrapping to all columns except CUSTTYPE and ACTIVECODE
for col in df_new.columns:
    if col not in ['CUSTTYPE', 'ACTIVECODE']:
        df_new[col] = df_new[col].astype(str).apply(quote_wrap)


# Add trailer row
trailer_row = pd.DataFrame([['TRAILER'] + [','] * (len(df_new.columns) - 2)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Define output file path
output_csv = os.path.join(os.path.dirname(file_path), 'STAGE_CUST_INFO.csv')

# Save to CSV
df_new.to_csv(output_csv, index=False)

# Confirmation message
print(f"CSV file saved at {output_csv}")
