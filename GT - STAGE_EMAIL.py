# GT - STAGE_EMAIL

import pandas as pd
import os

# File path
file_path = r"C:\Users\us85360\OneDrive - Grant Thornton LLP\0 - All Work\[c] 15 - Unitil\0_SM Work\BangorDataSAPSampleExtract.xlsx"

# Read the Excel file and load the specific sheet
df = pd.read_excel(file_path, sheet_name='ZCAMPAIGN', engine='openpyxl')

# Initialize df_new using relevant columns
df_new = pd.DataFrame()

# Assign CUSTOMERID and EMAILADDRESS with proper formatting
df_new['CUSTOMERID'] = df.iloc[:, 1].fillna('').astype(str).str.slice(0, 15)
df_new['EMAILADDRESS'] = df.iloc[:, 9].fillna('').astype(str).str.slice(0, 254)

# Add EMAILCODE and PRIORITY as empty columns for table structure
df_new['EMAILCODE'] = "1"  # Placeholder column
df_new['PRIORITY'] = ""   # Placeholder column

# Add additional columns with default empty values
additional_columns = ['UPDATEDATE']
for col in additional_columns:
    df_new[col] = " "

# Function to wrap values in double quotes, except for 'ADDRESSSEQ' (numeric) and 'UPDATEDATE' (date)
def quote_wrap(x):
    if pd.api.types.is_numeric_dtype(x) or isinstance(x, (int, float)):
        return x  # Keep numbers as they are
    return f'"{x}"' if pd.notna(x) else '""'  # Add double quotes, handle NaN

# Apply function to all columns except 'ADDRESSSEQ' and 'UPDATEDATE'
for col in df_new.columns:
    if col not in ['EMAILCODE', 'PRIORITY', 'UPDATEDATE']:
        df_new[col] = df_new[col].astype(str).apply(quote_wrap)




# Add trailer row
trailer_row = pd.DataFrame([['TRAILER'] + [','] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Define output file path
output_csv = os.path.join(os.path.dirname(file_path), 'STAGE_EMAIL.csv')

# Save to CSV
df_new.to_csv(output_csv, index=False)

# Confirmation message
print(f"CSV file saved at {output_csv}")
