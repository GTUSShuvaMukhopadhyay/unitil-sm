#CONV1 - STAGE_METERED_SVCS.py
# STAGE_METERED_SVCS.py

# we need to exclude the contractids in the list below from our data set ~ will code around it later

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
    if column_name in ['CUSTOMERID', 'LOCATIONID', 'METERNUMBER', 'INITIALSERVICEDATE', 'BILLINGSTARTDATE', 'LASTREADDATE', 'HHCOMMENTS', 'SERVICECOMMENTS', 'USERDEFINED', 'STOPESTIMATE', 'TAMPERCODE', 'UPDATEDATE', 'REMOVEDDATE']:
        return val  # Keep numeric values unquoted
    return '' if val in [None, 'nan', 'NaN', 'NAN'] else custom_quote(val)

# Define file paths
file_paths = {
    "ZDM_PREMDETAILS": "C:/Path/To/ZDM_PREMDETAILS.xlsx",
    "ZNC_ACTIVE_CUS": "C:/Path/To/ZNC_ACTIVE_CUS.xlsx",
    "EABL": "C:/Path/To/EABL 01012020 TO 2132025.xlsx",
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

# Extract CUSTOMERID from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["CUSTOMERID"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 9].fillna('').astype(str)

# Extract LOCATIONID from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["LOCATIONID"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 2].fillna('').astype(str)

# Extract METERNUMBER from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["METERNUMBER"] = data_sources["ZDM_PREMDETAILS"].iloc[:, 18].fillna('').astype(str)

# Define exclusion list for CUSTOMERID
excluded_customer_ids = {
    "210792305", "210806609", "210826823", "210800918", "210824447", "210830220", "210816965", 
    "200332427", "200611277", "210820685", "210793791", "200413813", "200437326", "200561498", 
    "210796711", "210797040", "210796579", "210796654", "210796769", "210796844", "210796909", "210796977"
}

# Define mappings
BILLINGRATE1_category_mapping = {
    "T_ME_RESID": "8002",
    "T_ME_SCISL": "8040",
    "T_ME_LCISL": "8042",
    "T_ME_SCITR": "8040",
    "T_ME_LCITR": "8042"
}

SALESCLASS1_category_mapping = {
    "T_ME_RESID": "8002",
    "T_ME_SCISL": "8040",
    "T_ME_LCISL": "8042",
    "T_ME_SCITR": "8240",
    "T_ME_LCITR": "8242"
}

BILLINGRATE2_category_mapping = {
    "T_ME_RESID": "8300",
    "T_ME_SCISL": "8302",
    "T_ME_LCISL": "8304",
    "T_ME_SCITR": "9800",
    "T_ME_LCITR": "9800"
}

SALESCLASS2_category_mapping = {
    "T_ME_RESID": "8002",
    "T_ME_SCISL": "8040",
    "T_ME_LCISL": "8042",
    "T_ME_SCITR": "8240",
    "T_ME_LCITR": "8242"
}

# Extract BILLINGRATE1, SALESCLASS1, BILLINGRATE2, and SALESCLASS2 from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    rate_category_column = data_sources["ZDM_PREMDETAILS"].iloc[:, 4].fillna('').astype(str)
    df_new["BILLINGRATE1"] = [BILLINGRATE1_category_mapping.get(rate_category_column[i], "") if df_new["CUSTOMERID"].iloc[i] not in excluded_customer_ids else "" for i in range(len(df_new))]
    df_new["SALESCLASS1"] = [SALESCLASS1_category_mapping.get(rate_category_column[i], "") if df_new["CUSTOMERID"].iloc[i] not in excluded_customer_ids else "" for i in range(len(df_new))]
    df_new["BILLINGRATE2"] = [BILLINGRATE2_category_mapping.get(rate_category_column[i], "") if df_new["CUSTOMERID"].iloc[i] not in excluded_customer_ids else "" for i in range(len(df_new))]
    df_new["SALESCLASS2"] = [SALESCLASS2_category_mapping.get(rate_category_column[i], "") if df_new["CUSTOMERID"].iloc[i] not in excluded_customer_ids else "" for i in range(len(df_new))]



# Extract INITIALSERVICEDATE and BILLINGSTARTDATE from ZNC_ACTIVE_CUS
if data_sources["ZNC_ACTIVE_CUS"] is not None:
    df_new["INITIALSERVICEDATE"] = pd.to_datetime(data_sources["ZNC_ACTIVE_CUS"].iloc[:, 7], errors='coerce').dt.strftime('%Y-%m-%d')
    df_new["BILLINGSTARTDATE"] = pd.to_datetime(data_sources["ZNC_ACTIVE_CUS"].iloc[:, 7], errors='coerce').dt.strftime('%Y-%m-%d')


# Extract LASTREADING and LASTREADDATE from ZNC_ACTIVE_CUS
if data_sources["EABL"] is not None:
    df_new["LASTREADING"] = pd.to_numeric(data_sources["EABL"].iloc[:, 8], errors='coerce')
    df_new["LASTREADDATE"] = pd.to_datetime(data_sources["EABL"].iloc[:, 8], errors='coerce')

# Extract MULTIPLIER from ZDM_PREMDETAILS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["MULTIPLIER"] = pd.to_numeric(data_sources["ZDM_PREMDETAILS"].iloc[:, 22], errors='coerce')

# Assign hardcoded values
df_new["APPLICATION"] = "5"
df_new["SERVICENUMBER"] = "1"
df_new["SERVICETYPE"] = "0"
df_new["METERREGISTER"] = "1"
df_new["SERVICESTATUS"] = "0"
df_new["LATITUDE"] = ""
df_new["READSEQUENCE"] = "0" 
df_new["LONGITUDE"] = ""
df_new["HHCOMMENTS"] = ""
df_new["SERVICECOMMENTS"] = ""
df_new["USERDEFINED"] = ""
df_new["STOPESTIMATE"] = ""
df_new["LOCATIONCODE"] = ""
df_new["INSTRUCTIONCODE"] = ""
df_new["TAMPERCODE"] = ""
df_new["AWCVALUE"] = ""
df_new["UPDATEDATE"] = ""
df_new["REMOVEDDATE"] = "" # NEED UPDATED MAPPING

# Extract INITIALSERVICEDATE and BILLINGSTARTDATE from ZNC_ACTIVE_CUS
if data_sources["ZDM_PREMDETAILS"] is not None:
    df_new["REMOVEDDATE"] = pd.to_datetime(data_sources["ZDM_PREMDETAILS"].iloc[:, 7], errors='coerce').dt.strftime('%Y-%m-%d')

# Drop duplicate records based on CUSTOMERID and LOCATIONID
df_new = df_new.drop_duplicates(subset=['CUSTOMERID', 'LOCATIONID'], keep='first')

# Add a trailer row with default values
trailer_row = pd.DataFrame([["TRAILER"] + [''] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)


# Reorder columns based on user preference
column_order = [
    "CUSTOMERID", "LOCATIONID", "APPLICATION", "SERVICENUMBER", "SERVICETYPE",
    "METERNUMBER", "METERREGISTER", "SERVICESTATUS", "INITIALSERVICEDATE",
    "BILLINGSTARTDATE", "BILLINGRATE1", "SALESCLASS1", "BILLINGRATE2",
    "SALESCLASS2", "READSEQUENCE", "LASTREADING", "LASTREADDATE", "MULTIPLIER",
    "LATITUDE", "LONGITUDE", "HHCOMMENTS", "SERVICECOMMENTS", "USERDEFINED",
    "STOPESTIMATE", "LOCATIONCODE", "INSTRUCTIONCODE", "TAMPERCODE", "AWCVALUE",
    "UPDATEDATE", "REMOVEDDATE"
]

df_new = df_new[column_order]

# Save to CSV
output_path = "C:/Path/To/STAGE_METERED_SVCS.csv"
df_new.to_csv(output_path, index=False, header=True, quoting=csv.QUOTE_NONE, escapechar='\\')

# Confirmation message
print(f"CSV file saved at {output_path}")
