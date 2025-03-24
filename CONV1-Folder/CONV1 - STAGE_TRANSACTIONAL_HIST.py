# CONV1 - STAGE_TRANSACTIONAL_HIST.py
# STAGE_TRANSACTIONAL_HIST.py

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
    "DFKKOP": r"C:\1DFKKOP01012025TO02282025.XLSX",
    "ZDM_PREMDETAILS": r"C:\ZDM_PREMDETAILS.XLSX",
    "EVER": r"C:\EVER.XLSX",
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

# Extract CUSTOMERID from DFKKOP (Column A = iloc[:, 0])
if data_sources["DFKKOP"] is not None:
    df_new["CUSTOMERID"] = data_sources["DFKKOP"].iloc[:, 0].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x)
    ).str.slice(0, 15)
    df_new["CUSTOMERID"] = df_new["CUSTOMERID"]  # for merging later

# --- Assign LOCATIONID based on Cont.Account -> Contract Account match ---
if data_sources["EVER"] is not None and data_sources["ZDM_PREMDETAILS"] is not None:
    ever_df = data_sources["EVER"]
    zdm_df = data_sources["ZDM_PREMDETAILS"]

    # Convert and clean keys for reliable merge
    ever_df["Cont.Account"] = ever_df["Cont.Account"].astype(str).str.strip()
    zdm_df["Contract Account"] = zdm_df["Contract Account"].apply(
        lambda x: str(int(x)) if pd.notna(x) else ""
    ).str.strip()

    # Build lookup table: Contract Account -> Premise (renamed to LOCATIONID)
    zdm_df["Premise"] = zdm_df["Premise"].apply(lambda x: str(int(x)) if pd.notna(x) else "")
    location_lookup = zdm_df[["Contract Account", "Premise"]].copy()
    location_lookup = location_lookup.rename(columns={"Premise": "LOCATIONID"})

    # Pull over Cont.Account into df_new from ever_df
    df_new["Cont.Account"] = ever_df["Cont.Account"]

    # Merge df_new with lookup
    df_new = df_new.merge(location_lookup, how="left", left_on="Cont.Account", right_on="Contract Account")

    # Drop helper column if needed
    df_new.drop(columns=["Contract Account"], inplace=True)



# Extract TRANSACTIONDATE from DFKKOP
if data_sources["DFKKOP"] is not None:
    df_new["TRANSACTIONDATE"] = pd.to_datetime(data_sources["DFKKOP"].iloc[:, 13], errors='coerce').dt.strftime('%Y-%m-%d')


# Extract BILLINGDATE from DFKKOP
if data_sources["DFKKOP"] is not None:
    df_new["BILLINGDATE"] = pd.to_datetime(data_sources["DFKKOP"].iloc[:, 11], errors='coerce').dt.strftime('%Y-%m-%d')


# Extract DUEDATE from DFKKOP
if data_sources["DFKKOP"] is not None:
    df_new["DUEDATE"] = pd.to_datetime(data_sources["DFKKOP"].iloc[:, 12], errors='coerce').dt.strftime('%Y-%m-%d')


# Extract BILLORINVOICENUMBER from DFKKOP
if data_sources["DFKKOP"] is not None:
    df_new["BILLORINVOICENUMBER"] = data_sources["DFKKOP"].iloc[:, 8].apply(
    lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x)
).str.strip()


# Extract TRANSACTIONAMOUNT from DFKKOP
if data_sources["DFKKOP"] is not None:
    df_new["TRANSACTIONAMOUNT"] = pd.to_numeric(data_sources["DFKKOP"].iloc[:, 6], errors='coerce').fillna(0)



# Assign hardcoded values
df_new["TAXYEAR"] = " "
df_new["TRANSACTIONTYPE"] = "" # issue is that DFKKOP  column E when you goto TFKHVO  does not provide a 2 digit numeric value on the mpping
df_new["TRANSACTIONDESCRIPTION"] = "" # issue is that DFKKOP  column E when you goto TFKHVO  does not provide a 2 digit numeric value on the mpping
df_new["APPLICATION"] = "5"
df_new["BILLTYPE"] = "" # issue is that DFKKOP  column E when you goto TFKHVO  does not provide a 2 digit numeric value on the mpping
df_new["TENDERTYPE"] = " "
df_new["UPDATEDATE"] = " "


# Function to wrap values in double quotes, but leave blanks and NaN as they are
def custom_quote(val):
    if pd.isna(val) or val == "" or val == " ":
        return ''
    return f'"{val}"'

def selective_custom_quote(val, column_name):
    if column_name in ['BILLORINVOICENUMBER','TRANSACTIONTYPE','TRANSACTIONAMOUNT','APPLICATION','BILLTYPE','TENDERTYPE']:
        return val
    return '' if val in [None, 'nan', 'NaN', 'NAN'] else custom_quote(val)

df_new = df_new.fillna('')
df_new = df_new.apply(lambda col: col.map(lambda x: selective_custom_quote(x, col.name)))


# Reorder columns
column_order = [
    "TAXYEAR", "CUSTOMERID", "LOCATIONID", "TRANSACTIONDATE", "BILLINGDATE",
    "DUEDATE", "BILLORINVOICENUMBER", "TRANSACTIONTYPE", "TRANSACTIONAMOUNT",
    "TRANSACTIONDESCRIPTION", "APPLICATION", "APPLICATION","BILLTYPE", "TENDERTYPE", "UPDATEDATE"
]
df_new = df_new[column_order]


# Add trailer row
trailer_row = pd.DataFrame([["TRAILER"] + [''] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Save to CSV
output_path = os.path.join(os.path.dirname(list(file_paths.values())[0]), 'STAGE_TRANSACTIONAL_HIST.csv')
df_new.to_csv(output_path, index=False, header=True, quoting=csv.QUOTE_NONE, escapechar='\\')
print(f"CSV file saved at {output_path}")
