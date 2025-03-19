# CONV1 - STAGE_BILLING_ACCT.py
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

# Define file paths
file_paths = {
    "ZDM_PREMDETAILS": r"C:\Users\us85360\Desktop\to delete\ZDM_PREMDETAILS.xlsx",
    "EVER": r"C:\Users\us85360\Desktop\to delete\EVER.xlsx",
    "DFKKOP": r"C:\Users\us85360\Desktop\to delete\DFKKOP.xlsx",
    "ZNC_ACTIVE_CUS": r"C:\Users\us85360\Desktop\to delete\ZNC_ACTIVE_CUS.xlsx",
    "DFKKCOH": r"C:\Users\us85360\Desktop\to delete\DFKKCOH.xlsx",
    "WRITE_OFF": r"C:\Users\us85360\Desktop\to delete\Write off customer history.XLSX",
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
    df_new["ACCOUNTNUMBER"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 9].fillna('').apply(lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x)).str.slice(0, 15)

# Extract CUSTOMERID from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["CUSTOMERID"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 7].fillna('').apply(lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x)).str.slice(0, 15)

# Extract LOCATIONID from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["LOCATIONID"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 2].fillna('').astype(str)

# Assign hardcoded values
df_new["STATUSCODE"] = "0"
df_new["ADDRESSSEQ"] = "1"
df_new["TAXCODE"] = "0"
df_new["ARCODE"] = "8"
df_new["BANKCODE"] = "8"
df_new["DWELLINGUNITS"] = "1"
df_new["STOPSHUTOFF"] = "0"
df_new["STOPPENALTY"] = "0"
df_new["DUEDATE"] = " "  # Data doesn't map
df_new["SICCODE"] = " "
df_new["BUNCHCODE"] = " "
df_new["SHUTOFFDATE"] = " "
df_new["PIN"] = " "
df_new["DEFERREDDUEDATE"] = " "
df_new["LASTNOTICECODE"] = "0"
df_new["LASTNOTICEDATE"] = " "
df_new["CASHONLY"] = " "
df_new["NEMLASTTRUEUPDATE"] = " "
df_new["NEMNEXTTRUEUPDATE"] = " "
df_new["ENGINEERNUM"] = " "
df_new["UPDATEDATE"] = " "

# Extract OPENDATE from EVER and ensure it is formatted as YYYY-MM-DD
if data_sources["EVER"] is not None:
    ever_data = data_sources["EVER"]
    ever_data["Cont.Account"] = ever_data.iloc[:, 79].astype(str)
    ever_data["M/I Date"] = pd.to_datetime(ever_data.iloc[:, 83], errors='coerce').dt.strftime('%Y-%m-%d')
    df_new = df_new.merge(ever_data[["Cont.Account", "M/I Date"]], left_on="ACCOUNTNUMBER", right_on="Cont.Account", how="left")
    df_new.rename(columns={"M/I Date": "OPENDATE"}, inplace=True)
    df_new.drop(columns=["Cont.Account"], inplace=True)

# Extract TERMINATEDDATE from EVER and ensure it is formatted as YYYY-MM-DD
if data_sources["EVER"] is not None:
    ever_data["M/O Date"] = pd.to_datetime(ever_data.iloc[:, 84], errors='coerce').dt.strftime('%Y-%m-%d')
    df_new = df_new.merge(ever_data[["Cont.Account", "M/O Date"]], left_on="ACCOUNTNUMBER", right_on="Cont.Account", how="left")
    df_new.rename(columns={"M/O Date": "TERMINATEDDATE"}, inplace=True)
    df_new.drop(columns=["Cont.Account"], inplace=True)  # Drop after merge


# Ensure Write off customer history's Cont.Account (Column 1 - B) retains leading zeros
if data_sources["WRITE_OFF"] is not None:
    write_off_accounts = data_sources["WRITE_OFF"].iloc[:, 1].astype(str).unique()
else:
    write_off_accounts = set()

# Function to assign ACTIVECODE based on the corrected logic
def assign_active_code_corrected(account_number):
    matched_rows = ever_data[ever_data["Cont.Account"] == account_number]
    
    if not matched_rows.empty:
        mo_date = matched_rows.iloc[0]["M/O Date"]  # Fetch M/O Date
        
        if mo_date == "12/31/9999":
            return "0"
        elif account_number in write_off_accounts:
            return "4"
        else:
            return "2"
    
    return ""  # Default empty if no match found

# Apply function to determine ACTIVECODE
df_new["ACTIVECODE"] = df_new["ACCOUNTNUMBER"].apply(assign_active_code_corrected)

# Extract PENALTYCODE from ZNC_ACTIVE_CUS
if data_sources["ZNC_ACTIVE_CUS"] is not None:
    penalty_data = data_sources["ZNC_ACTIVE_CUS"]
    penalty_data["Cont.Account"] = penalty_data.iloc[:, 3].astype(str)  # Column D
    penalty_data["Fact grp"] = penalty_data.iloc[:, 22].astype(str)  # Column W
    
    penalty_mapping = {"RES": "53", "LCI": "55", "LCIT": "55", "SCI": "55", "SCIT": "55"}
    
    df_new = df_new.merge(penalty_data[["Cont.Account", "Fact grp"]], left_on="CUSTOMERID", right_on="Cont.Account", how="left")
    df_new["PENALTYCODE"] = df_new["Fact grp"].map(penalty_mapping).fillna("")
    df_new.drop(columns=["Cont.Account", "Fact grp"], inplace=True)



# Extract TAXTYPE from ZNC_ACTIVE_CUS
if data_sources["ZNC_ACTIVE_CUS"] is not None:
    tax_mapping = {"RES": "0", "LCI": "1", "LCIT": "1", "SCI": "1", "SCIT": "1"}
    df_new = df_new.merge(penalty_data[["Cont.Account", "Fact grp"]], left_on="CUSTOMERID", right_on="Cont.Account", how="left")
    df_new["TAXTYPE"] = df_new["Fact grp"].map(tax_mapping).fillna("")
    df_new.drop(columns=["Cont.Account", "Fact grp"], inplace=True)



# Function to wrap values in double quotes, but leave blanks and NaN as they are
def custom_quote(val):
    """Wraps all values in quotes except for blank or NaN ones."""
    # If the value is NaN, None, or blank, leave it empty
    if pd.isna(val) or val == "" or val == " ":
        return ''  # Return an empty string for NaN or blank fields
    return f'"{val}"'  # Wrap other values in double quotes
 
# Apply custom_quote function to all columns
df_new = df_new.fillna('')
 
def selective_custom_quote(val, column_name):
    if column_name in ['ACTIVECODE','STATUSCODE','ADDRESSSEQ','PENALTYCODE','TAXCODE','TAXTYPE','ARCODE','BANKCODE','DWELLINGUNITS','STOPSHUTOFF','STOPPENALTY']:
        return val  # Keep numeric values unquoted
    return '' if val in [None, 'nan', 'NaN', 'NAN'] else custom_quote(val)
 
df_new = df_new.apply(lambda col: col.map(lambda x: selective_custom_quote(x, col.name)))

# Remove any records missing ACCOUNTNUMBER
df_new = df_new[df_new['ACCOUNTNUMBER'].notna() & (df_new['ACCOUNTNUMBER'] != '')]

# Drop duplicate records based on ACCOUNTNUMBER, CUSTOMERID, and LOCATIONID
df_new = df_new.drop_duplicates(subset=['ACCOUNTNUMBER','CUSTOMERID','LOCATIONID'], keep='first')

# Reorder columns based on user preference
column_order = [
    "ACCOUNTNUMBER", "CUSTOMERID", "LOCATIONID", "ACTIVECODE", "STATUSCODE", "ADDRESSSEQ", "PENALTYCODE", "TAXCODE", "TAXTYPE", "ARCODE", "BANKCODE", "OPENDATE", "TERMINATEDDATE", "DWELLINGUNITS", "STOPSHUTOFF", "STOPPENALTY", "DUEDATE", "SICCODE", "BUNCHCODE", "SHUTOFFDATE", "PIN", "DEFERREDDUEDATE", "LASTNOTICECODE", "LASTNOTICEDATE", "CASHONLY", "NEMLASTTRUEUPDATE", "NEMNEXTTRUEUPDATE", "ENGINEERNUM", "UPDATEDATE"
]
df_new = df_new[column_order]


# Add a trailer row with default values
trailer_row = pd.DataFrame([["TRAILER"] + [''] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)





# Define output path for the CSV file
output_path = os.path.join(os.path.dirname(list(file_paths.values())[0]), 'STAGE_BILLING_ACCT.csv')
 
# Save to CSV with proper quoting and escape character
df_new.to_csv(output_path, index=False, header=True, quoting=csv.QUOTE_NONE, escapechar='\\')
 
# Confirmation message
print(f"CSV file saved at {output_path}")
