# CONV1 - STAGE_CONSUMPTION_HIST.py
# STAGE_CONSUMPTION_HIST.py

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
    "ZMECON": r"C:\",
    "ZDM_PREMDETAILS": r"C:\",
    "EABL": r"C:\",
    "CD": r"C:\codes and descriptions.xl"
    "MM": r"C:\METERMULTIPLIER_PressureFactor.xlsx",
    "TF": r"C:\ThermFactor.xlsx",
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

# Extract CUSTOMERID from ZMECON (Column A = iloc[:, 0])
if data_sources["ZMECON"] is not None:
    df_new["CUSTOMERID"] = data_sources["ZMECON"].iloc[:, 0].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x)
    ).str.slice(0, 15)
    df_new["CUSTOMERID"] = df_new["CUSTOMERID"]  # for merging 
    

    
# Extract LOCATIONID from ZMECON
if data_sources["ZMECON"] is not None:
    df_new["LOCATIONID"] = data_sources["ZMECON"].iloc[:, 25].fillna('').astype(str)

# Extract METERNUMBER from ZMECON
if data_sources["ZMECON"] is not None:
    df_new["METERNUMBER"] = data_sources["ZMECON"].iloc[:, 20].fillna('').astype(str)

# Extract CURRREADDATE from ZMECON
if data_sources["ZMECON"] is not None:
    df_new["CURRREADDATE"] = pd.to_datetime(data_sources["ZMECON"].iloc[:, 23], errors='coerce').dt.strftime('%Y-%m-%d')

# Map METERMULTIPLIER from MM based on METERNUMBER
if data_sources["MM"] is not None and "Meter #1" in data_sources["MM"].columns and "PressureFactor" in data_sources["MM"].columns:
    mm_df = data_sources["MM"]
    
    # Standardize both sides to string
    mm_df["Meter #1"] = mm_df["Meter #1"].astype(str)
    df_new["METERNUMBER"] = df_new["METERNUMBER"].astype(str)

    # Create the mapping dictionary
    meter_to_multiplier = dict(zip(mm_df["Meter #1"], mm_df["PressureFactor"]))

    # Apply mapping to METERMULTIPLIER column
    df_new["METERMULTIPLIER"] = df_new["METERNUMBER"].map(meter_to_multiplier)
else:
    print("Warning: MM file missing 'Meter #1' and/or 'PressureFactor' columns.")



# Extract PREVREADDATE from ZMECON
if data_sources["ZMECON"] is not None:
    df_new["PREVREADDATE"] = pd.to_datetime(data_sources["ZMECON"].iloc[:, 22], errors='coerce').dt.strftime('%Y-%m-%d')


# Extract CURRREADING from EABL
if data_sources["EABL"] is not None:
    df_new["CURRREADING"] = pd.to_numeric(data_sources["EABL"].iloc[:, 8], errors='coerce').fillna(0)

# Extract RAWUSAGE from EABL
if data_sources["EABL"] is not None:
    df_new["RAWUSAGE"] = pd.to_numeric(data_sources["EABL"].iloc[:, 8], errors='coerce').fillna(0)


# Extract BILLINGUSAGE from ZMECON
if data_sources["ZMECON"] is not None:
    df_new["BILLINGUSAGE"] = pd.to_numeric(data_sources["ZMECON"].iloc[:, 21], errors='coerce').fillna(0)


# Extract BILLEDDATE from ZMECON
if data_sources["ZMECON"] is not None:
    df_new["BILLEDDATE"] = pd.to_datetime(data_sources["ZMECON"].iloc[:, 23], errors='coerce').dt.strftime('%Y-%m-%d')

# -------------------------------
# Assign THERMFACTOR from ThermFactor.xlsx
# -------------------------------

# Strip whitespace from TF column names
data_sources["TF"].columns = data_sources["TF"].columns.str.strip()

# Convert Valid from / to in TF to datetime
therm_df = data_sources["TF"]
therm_df["Valid from"] = pd.to_datetime(therm_df["Valid from"], errors="coerce")
therm_df["Valid to"] = pd.to_datetime(therm_df["Valid to"], errors="coerce")

# Ensure your df_new has date columns to compare — adjust as needed:
df_new["DATE_FROM"] = pd.to_datetime(data_sources["ZMECON"].iloc[:, 22], errors='coerce')  # example: PREVREADDATE
df_new["DATE_TO"] = pd.to_datetime(data_sources["ZMECON"].iloc[:, 23], errors='coerce')   # example: CURRREADDATE

# Function to match and pull Avg. BTU
def find_matching_btu(start, end):
    match = therm_df[
        (therm_df["Valid from"] <= end) & (therm_df["Valid to"] >= start)
    ]
    if not match.empty:
        return match.iloc[0]["Avg. BTU"]
    return None  # Or return default like 1.0 if no match

# Apply to df_new
df_new["THERMFACTOR"] = df_new.apply(
    lambda row: find_matching_btu(row["DATE_FROM"], row["DATE_TO"]), axis=1
)



# Assign hardcoded values
df_new["APPLICATION"] = "5"
df_new["SERVICENUMBER"] = "1"
df_new["METERREGISTER"] = "1"
df_new["READINGCODE"] = " " # major issues with field Need to review an example
df_new["READINGTYPE"] = " " # major issues with field Need to review an example
df_new["PREVREADING"] = " " # major issues with field Need to review an example
df_new["UNITOFMEASURE"] = "CF"
df_new["READERID"] = " " 
df_new["BILLEDAMOUNT"] = " " 
df_new["BILLINGBATCHNUMBER"] = " " 
df_new["BILLINGRATE"] = " " 
df_new["SALESREVENUECLASS"] = " " 
df_new["HEATINGDEGREEDAYS"] = " " 
df_new["COOLINGDEGREEDAYS"] = " " 
df_new["UPDATEDATE"] = " " 




# Function to wrap values in double quotes, but leave blanks and NaN as they are
def custom_quote(val):
    if pd.isna(val) or val == "" or val == " ":
        return ''
    return f'"{val}"'

def selective_custom_quote(val, column_name):
    if column_name in ['APPLICATION','SERVICENUMBER','METERREGISTER','READINGCODE','READINGTYPE','CURRREADING','PREVREADING','RAWUSAGE','BILLINGUSAGE','METERMULTIPLIER','THERMFACTOR','BILLEDAMOUNT','BILLINGBATCHNUMBER','BILLINGRATE','SALESREVENUECLASS','HEATINGDEGREEDAYS','COOLINGDEGREEDAYS']:
        return val
    return '' if val in [None, 'nan', 'NaN', 'NAN'] else custom_quote(val)

df_new = df_new.fillna('')
df_new = df_new.apply(lambda col: col.map(lambda x: selective_custom_quote(x, col.name)))

# Reorder columns
column_order = [
    "CUSTOMERID", "LOCATIONID", "APPLICATION", "SERVICENUMBER", "METERNUMBER",
    "METERREGISTER", "READINGCODE", "READINGTYPE", "CURRREADDATE",
    "PREVREADDATE", "CURRREADING", "PREVREADING","UNITOFMEASURE","RAWUSAGE","BILLINGUSAGE","METERMULTIPLIER",
    "BILLEDDATE","THERMFACTOR","READERID","BILLEDAMOUNT","BILLINGBATCHNUMBER","BILLINGRATE",
    "SALESREVENUECLASS", "HEATINGDEGREEDAYS", "COOLINGDEGREEDAYS","UPDATEDATE"
]
df_new = df_new[column_order]

# Add trailer row
trailer_row = pd.DataFrame([["TRAILER"] + [''] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Save to CSV
output_path = os.path.join(os.path.dirname(list(file_paths.values())[0]), 'STAGE_CONSUMPTION_HIST.csv')
df_new.to_csv(output_path, index=False, header=True, quoting=csv.QUOTE_NONE, escapechar='\\')
print(f"CSV file saved at {output_path}")
