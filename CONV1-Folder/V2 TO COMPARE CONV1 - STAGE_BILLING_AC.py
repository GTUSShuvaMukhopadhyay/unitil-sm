# V2 TO COMPARE CONV1 - STAGE_BILLING_ACCT.py
# STAGE_BILLING_ACCT.py

import pandas as pd
import os
import re
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

# Function to wrap values in double quotes, but leave blanks and NaN as they are
def custom_quote(val):
    """Wraps all values in quotes except for blank or NaN ones."""
    if pd.isna(val) or val == "" or val == " ":
        return ''  # Return an empty string for NaN or blank fields
    return f'"{val}"'  # Wrap other values in double quotes

# Apply selective quoting
def selective_custom_quote(val, column_name):
    if column_name in ['ACTIVECODE', 'STATUSCODE', 'ADDRESSSEQ', 'PENALTYCODE', 'TAXCODE', 'TAXTYPE', 'ARCODE', 'BANKCODE', 'DWELLINGUNITS', 'STOPSHUTOFF', 'STOPPENALTY', 'LASTNOTICECODE']:
        return val  # Keep numeric values unquoted
    return '' if val in [None, 'nan', 'NaN', 'NAN'] else custom_quote(val)

# Define file paths
file_paths = {
    "ZDM_PREMDETAILS": "C:/Path/To/ZDM_PREMDETAILS.xlsx",
    "EVER": "C:/Path/To/EVER.xlsx",
    "DFKKOP": "C:/Path/To/DFKKOP.xlsx",
    "ZNC_ACTIVE_CUS": "C:/Path/To/ZNC_ACTIVE_CUS.xlsx",
    "DFKKCOH": "C:/Path/To/DFKKCOH.xlsx",
    "WRITE_OFF": "C:/Path/To/Write off customer history.XLSX",
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

# Extract ACCOUNTNUMBER from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["ACCOUNTNUMBER"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 6].fillna('').astype(str)

# Extract CUSTOMERID from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["CUSTOMERID"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 4].fillna('').astype(str)

# Extract LOCATIONID from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["LOCATIONID"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 12].fillna('').astype(str)

# Assign default empty set for WRITE_OFF accounts
write_off_accounts = set()
if data_sources["WRITE_OFF"] is not None:
    write_off_accounts = set(data_sources["WRITE_OFF"].iloc[:, 1].astype(str).dropna())

# Restore ACTIVECODE logic
if data_sources["EVER"] is not None:
    ever_data = data_sources["EVER"]
    def assign_active_code(account_number):
        matched_row = ever_data.loc[ever_data.iloc[:, 79] == account_number]
        if not matched_row.empty:
            mo_date = matched_row.iloc[0, 84]  # M/O Date
            if mo_date == "12/31/9999":
                return "0"
            elif account_number in write_off_accounts:
                return "4"
            else:
                return "2"
        return ""
    df_new["ACTIVECODE"] = df_new["ACCOUNTNUMBER"].apply(assign_active_code)

# Restore PENALTYCODE and TAXTYPE logic
if data_sources["ZNC_ACTIVE_CUS"] is not None:
    active_cus_data = data_sources["ZNC_ACTIVE_CUS"]
    account_mapping = active_cus_data.set_index(active_cus_data.iloc[:, 3].astype(str))[23].to_dict()
    def assign_penalty_code(account_number):
        fact_grp = account_mapping.get(account_number, "")
        if fact_grp == "RES":
            return "53"
        elif fact_grp in ["LCI", "LCIT", "SCI", "SCIT"]:
            return "55"
        return ""
    def assign_tax_type(account_number):
        fact_grp = account_mapping.get(account_number, "")
        if fact_grp == "RES":
            return "0"
        elif fact_grp in ["LCI", "LCIT", "SCI", "SCIT"]:
            return "1"
        return ""
    df_new["PENALTYCODE"] = df_new["ACCOUNTNUMBER"].apply(assign_penalty_code)
    df_new["TAXTYPE"] = df_new["ACCOUNTNUMBER"].apply(assign_tax_type)

# Drop duplicate records based on ACCOUNTNUMBER, CUSTOMERID, and LOCATIONID
df_new = df_new.drop_duplicates(subset=['ACCOUNTNUMBER', 'CUSTOMERID', 'LOCATIONID'], keep='first')

# Save to CSV
output_path = "C:/Path/To/STAGE_BILLING_ACCT.csv"
df_new.to_csv(output_path, index=False, header=True, quoting=csv.QUOTE_NONE, escapechar='\\')

# Confirmation message
print(f"CSV file saved at {output_path}")
