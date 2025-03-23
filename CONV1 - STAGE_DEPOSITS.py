# CONV1 - STAGE_DEPOSITS.py
# STAGE_DEPOSITS.py

import pandas as pd
import os
import csv  # Import the correct CSV module

# CSV Staging File Checklist
CHECKLIST = [
    "✅ Filename must match the entry in Column D of the All Tables tab.",
    "✅ Filename must be in uppercase except for '.csv' extension.",
    "✅ The first record in the file must be the header row.",
    "✅ Ensure no extraneous rows (including blank rows) are present in the file.",
    "✅ All non-numeric fields must be enclosed in double quotes.",
    "✅ The last row in the file must be 'TRAILER' followed by commas.",
    "✅ Replace all CRLF (X'0d0a') in customer notes with ~^[",
    "✅ Ensure all dates are in 'YYYY-MM-DD' format.",
]

def print_checklist():
    print("CSV Staging File Validation Checklist:")
    for item in CHECKLIST:
        print(item)

print_checklist()

# Define file paths
file_paths = {
    "FPD2": r"C:\",
    "ZDM_PREMDETAILS": r"C:\",
    "ZMECON": r"C:\",
    "DFKKOP": r"C:\",
}

# Load the data from each spreadsheet
data_sources = {}
for name, path in file_paths.items():
    try:
        data_sources[name] = pd.read_excel(path, sheet_name="Sheet1", engine="openpyxl")
    except Exception as e:
        data_sources[name] = None
        print(f"Error loading {name}: {e}")

# Initialize df_new as an empty DataFrame
df_new = pd.DataFrame()

# Extract CUSTOMERID from FPD2 (Column L = iloc[:, 11])
if data_sources["FPD2"] is not None:
    df_new["CUSTOMERID"] = data_sources["FPD2"].iloc[:, 11].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x)
    ).str.slice(0, 15)
    df_new["Contract Account"] = df_new["CUSTOMERID"]  # for merging later

# Extract LOCATIONID from ZDM_PREMDETAILS (Column C), match on Contract Account (FPD2 col L and ZDM col J)
if data_sources["FPD2"] is not None and data_sources["ZDM_PREMDETAILS"] is not None:
    # Clean ZDM_PREMDETAILS contract accounts (col J = iloc[:, 9])
    zdm_ca = data_sources["ZDM_PREMDETAILS"].iloc[:, 9].astype(str).str.lstrip('0').str.strip()

    # Get Premise from col C = iloc[:, 2]
    zdm_temp = pd.DataFrame({
        "Contract Account": zdm_ca,
        "LOCATIONID": data_sources["ZDM_PREMDETAILS"].iloc[:, 2]
    })

    # Merge to get LOCATIONID
    df_new = df_new.merge(zdm_temp, on="Contract Account", how="left")

# Create DEPOSITSTATUS based on 'Description of Security Deposit Status' (Column K in FPD2)
if data_sources["FPD2"] is not None:
    df_new["DEPOSITSTATUS"] = data_sources["FPD2"].iloc[:, 10].apply(lambda x: 2 if x == "Paid" else (90 if x == "Request" else 1))

# Extract DEPOSITDATE from FPD2
if data_sources["FPD2"] is not None:
    df_new["DEPOSITDATE"] = pd.to_datetime(data_sources["FPD2"].iloc[:, 4], errors='coerce').dt.strftime('%Y-%m-%d')

# Extract DEPOSITAMOUNT from FPD2
if data_sources["FPD2"] is not None:
    df_new["DEPOSITAMOUNT"] = pd.to_numeric(data_sources["FPD2"].iloc[:, 8], errors='coerce').fillna(0)

# Extract DEPOSITINTERESTCALCDATE from FPD2
if data_sources["FPD2"] is not None:
    df_new["DEPOSITINTERESTCALCDATE"] = pd.to_datetime(data_sources["FPD2"].iloc[:, 4], errors='coerce').dt.strftime('%Y-%m-%d')

# Get max date for DEPOSITINTERESTCALCDATE from 'Pstng Date' (Column L in FPD2)
if data_sources["FPD2"] is not None:
    data_sources["FPD2"]["Pstng Date"] = pd.to_datetime(data_sources["FPD2"].iloc[:, 11], errors='coerce')
    max_date = data_sources["FPD2"]["Pstng Date"].max()
    df_new["DEPOSITINTERESTCALCDATE"] = max_date.strftime('%Y-%m-%d') if pd.notna(max_date) else ''


# Calculate the difference in months between 'End Date' (Column F) and Today
if data_sources["FPD2"] is not None:
    today = pd.Timestamp.today()
    data_sources["FPD2"]["End Date"] = pd.to_datetime(data_sources["FPD2"].iloc[:, 5], errors='coerce')
    df_new["DEPOSITREFUNDMONTHS"] = (data_sources["FPD2"]["End Date"].dt.year - today.year) * 12 + (data_sources["FPD2"]["End Date"].dt.month - today.month)

# Assign hardcoded values
df_new["APPLICATION"] = "25"
df_new["DEPOSITKIND"] = "CASH"
df_new["DEPOSITBILLEDFLAG"] = " "
df_new["DEPOSITACCRUEDINTEREST"] = " "
df_new["UPDATEDATE"] = " "

# Function to wrap values in double quotes, but leave blanks and NaN as they are
def custom_quote(val):
    if pd.isna(val) or val == "" or val == " ":
        return ''
    return f'"{val}"'

def selective_custom_quote(val, column_name):
    if column_name in ['APPLICATION','DEPOSITSTATUS','DEPOSITKIND','DEPOSITAMOUNT','DEPOSITACCRUEDINTEREST','DEPOSITREFUNDMONTHS']:
        return val
    return '' if val in [None, 'nan', 'NaN', 'NAN'] else custom_quote(val)

df_new = df_new.fillna('')
df_new = df_new.apply(lambda col: col.map(lambda x: selective_custom_quote(x, col.name)))

# Reorder columns
column_order = [
    "CUSTOMERID", "LOCATIONID", "APPLICATION", "DEPOSITSTATUS", "DEPOSITKIND",
    "DEPOSITDATE", "DEPOSITAMOUNT", "DEPOSITBILLEDFLAG", "DEPOSITACCRUEDINTEREST",
    "DEPOSITINTERESTCALCDATE", "DEPOSITREFUNDMONTHS", "UPDATEDATE"
]
df_new = df_new[column_order]

# Add trailer row
trailer_row = pd.DataFrame([["TRAILER"] + [''] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Save to CSV
output_path = os.path.join(os.path.dirname(list(file_paths.values())[0]), 'STAGE_DEPOSITS.csv')
df_new.to_csv(output_path, index=False, header=True, quoting=csv.QUOTE_NONE, escapechar='\\')
print(f"CSV file saved at {output_path}")
