import pandas as pd
import re
import csv

# Load the data
file_path = r"C:\Users\US97684\Downloads\documents_20250317_New\Outbound\ZDM_PREMDETAILS.XLSX"
file_path2 = r"C:\Users\US97684\Downloads\documents_20250317_New\Outbound\TE422.XLSX"
file_path1 = r"C:\Users\US97684\Downloads\documents_20250317_New\Premise_clean_final.xlsx"
df = pd.read_excel(file_path, sheet_name='Sheet1', engine='openpyxl')
df_Portion = pd.read_excel(file_path2, sheet_name='Sheet1', engine='openpyxl')
df_Premise = pd.read_excel(file_path1, sheet_name='Clean_Data', engine='openpyxl')

# Load configuration file for suffix lookup
config_path = r"C:\Users\US97684\Downloads\documents_20250317_New\Configuration.xlsx"
sheet1 = pd.read_excel(config_path, sheet_name='Street Abbreviation', engine='openpyxl')
sheet2 = pd.read_excel(config_path, sheet_name='Premise Designation', engine='openpyxl')

# Initialize df_new as an empty DataFrame
df_new = pd.DataFrame()

# Column 1: Column B (index 1)
df_new['LOCATIONID'] = df.iloc[:, 2]

# Apply the function to extract street number from the relevant column (iloc(26))
def fetch_streetnumber(location_id):
    location_id = str(location_id).strip()  # Convert LOCATIONID to string and remove extra spaces
    premise_clean = df_Premise.iloc[:, 0].astype(str).str.strip()  # Clean whitespace and convert to string  
    matched_row = df_Premise[premise_clean.str.contains(location_id)]
    
    if not matched_row.empty:
        return str(matched_row.iloc[0, 3]).strip()  # Return street number from the second column
    else:
        return None

df_new['STREETNUMBER'] = df_new['LOCATIONID'].apply(fetch_streetnumber)

# Move suffix to STREETNUMBERSUFFIX if STREETNUMBER contains letters
def move_suffix_to_streetnumbersuffix(streetnumber):
    if streetnumber:
        match = re.match(r'(\d+)([A-Za-z]+)', streetnumber)  # Match numbers followed by letters
        if match:
            # If match is found, separate street number and suffix
            return match.group(1), match.group(2)
    return streetnumber, ""  # Return original street number if no suffix

# Apply the function to split STREETNUMBER and assign suffix to STREETNUMBERSUFFIX
df_new[['STREETNUMBER', 'STREETNUMBERSUFFIX']] = df_new['STREETNUMBER'].apply(
    lambda x: pd.Series(move_suffix_to_streetnumbersuffix(x))
)

df_new['STREETNAME'] = ""

# STREETNAME
def fetch_streetname(location_id):
    location_id = str(location_id).strip()  # Convert LOCATIONID to string and remove extra spaces
    premise_clean = df_Premise.iloc[:, 0].astype(str).str.strip()  # Clean whitespace and convert to string  

    # Check if location_id is found in premise_clean using case-insensitive matching
    matched_row = df_Premise[premise_clean.str.contains(location_id, case=False, na=False)]  # Added case=False and na=False
    
    if not matched_row.empty:
        street_name = str(matched_row.iloc[0, 5]).strip()  # Assuming street name is in the 6th column
        return street_name if street_name != 'nan' else ""  # Return empty string if nan is found
    else:
        return ""  # Return empty string if no match is found

df_new['STREETNAME'] = df_new['LOCATIONID'].apply(fetch_streetname)

def fetch_designation(location_id):
    location_id = str(location_id).strip()  # Convert LOCATIONID to string and remove extra spaces
    premise_clean = df_Premise.iloc[:, 0].astype(str).str.strip()  # Clean whitespace and convert to string  
    matched_row = df_Premise[premise_clean.str.contains(location_id)]
    
    if not matched_row.empty:
        designation = str(matched_row.iloc[0, 8]).strip()
        designation = designation.replace("-", "").replace(".", "")
        return designation
    else:
        return None
df_new['DESIGNATION'] = df_new['LOCATIONID'].apply(fetch_designation)

df_new['ADDITIONALDESC'] = ""  # Initially leave ADDITIONALDESC empty

# Town
def fetch_town(location_id):
    location_id = str(location_id).strip()  # Convert LOCATIONID to string and remove extra spaces
    premise_clean = df_Premise.iloc[:, 0].astype(str).str.strip()  # Clean whitespace and convert to string  
    matched_row = df_Premise[premise_clean.str.contains(location_id)]
    
    if not matched_row.empty:
        return str(matched_row.iloc[0, 2]).strip()  # Return street name from the second column
    else:
        return None
df_new['TOWN'] = df_new['LOCATIONID'].apply(fetch_town)

df_new['STATE'] = "ME"

# Ensure ZIPCODE is always treated as a string with leading zeros
df_new['ZIPCODE'] = df.iloc[:, 27].astype(str).str.zfill(5)  # Ensure it's 5 digits, filling leading zeros if necessary


# Handle ZIPPLUSFOUR and any invalid ZIPCODE entries
ZIPCODE = pd.to_numeric(df_new['ZIPCODE'], errors='coerce')  # Convert ZIPCODE to numeric, errors coerced to NaN

# Ensure proper handling of NaN or invalid ZIPCODE values when adding '4'
df_new['ZIPPLUSFOUR'] = ZIPCODE.apply(
    lambda x: str(int(x) + 4) if pd.notna(x) and x != 0 else '00000'  # Handle NaN or zero cases
)

df_new['OWNERCUSTOMERID'] = ""
df_new['OWNERMAILSEQ'] = "1"
# Add PROPERTYCLASS logic
def map_property_class(value):
    mapping = {
        'G_ME_RESID': 1,
        'T_ME_RESID': 1,
        'G_ME_SCISL': 2,
        'T_ME_SCISL': 2,
        'T_ME_LIHEA': 1,
        'G_ME_LCISL': 2,
        'T_ME_LCISL': 2,
        'T_ME_LCITR': 2,
        'T_ME_SCITR': 2
    }
    return mapping.get(value, 1)  # Default to "1" if value not found in the mapping

# Apply the function to the df_new['PROPERTYCLASS'] column based on ZDM_PREMDETAILS column E
df_new['PROPERTYCLASS'] = df.iloc[:, 4].apply(map_property_class)
df_new['TAXDISTRICT'] = 8

# Create the dictionary for BILLINGCYCLE and READINGROUTE mapping
billing_and_reading_map = {
    "MEOTP01": 801,
    "MEOTP02": 802,
    "MEOROP01": 803,
    "MEOROP02": 804,
    "MEOROP03": 805,
    "MEBGRP01": 806,
    "MEBGRP02": 807,
    "MEBGRP03": 808,
    "MEBGRP04": 809,
    "MEBGRP05": 810,
    "MEBGRP06": 811,
    "MEBGRP07": 812,
    "MEBGRP08": 813,
    "MEBGRP09": 814,
    "MEBRWP01": 815,
    "MEBRWP02": 816,
    "MEBRWP03": 817,
    "MEBCKP01": 819,
    "MELINC01": 820,
    "METRNP01": 822
}

# Apply the mapping logic to both the BILLINGCYCLE and READINGROUTE columns
def map_billing_and_reading(location_id):
    location_id = str(location_id).strip()  # Convert LOCATIONID to string and remove extra spaces
    # Return the mapped value for either BILLINGCYCLE or READINGROUTE from the dictionary
    return billing_and_reading_map.get(location_id, "")
df_new['BILLINGCYCLE'] = ""  # Initially set to empty, will map below
df_new['READINGROUTE'] = ""  # Initially set to empty, will map below

# Apply the function to both columns
df_new['BILLINGCYCLE'] = df.iloc[:, 0].apply(map_billing_and_reading)
df_new['READINGROUTE'] = df.iloc[:, 0].apply(map_billing_and_reading)

df_new['SERVICEAREA'] = "80"
df_new['SERVICEFACILITY'] = ""
df_new['PRESSUREDISTRICT'] = ""
df_new['LATITUDE'] = ""
df_new['LONGITUDE'] = ""
df_new['MAPNUMBER'] = ""
df_new['PARCELID'] = ""
df_new['PARCELAREATYPE'] = ""
df_new['PARCELAREA'] = ""
df_new['IMPERVIOUSSQUAREFEET'] = ""
df_new['SUBDIVISION'] = ""
df_new['GISID'] = ""
df_new['FOLIOSEGMENT1'] = ""
df_new['FOLIOSEGMENT2'] = ""
df_new['FOLIOSEGMENT3'] = ""
df_new['FOLIOSEGMENT4'] = ""
df_new['FOLIOSEGMENT5'] = ""
df_new['PROPERTYUSECLASSIFICATION1'] = ""
df_new['PROPERTYUSECLASSIFICATION2'] = ""
df_new['PROPERTYUSECLASSIFICATION3'] = ""
df_new['PROPERTYUSECLASSIFICATION4'] = ""
df_new['PROPERTYUSECLASSIFICATION5'] = ""
df_new['UPDATEDATE'] = ""

# Remove rows with missing required fields
required_columns = [
    'LOCATIONID', 'STREETNAME', 'TOWN', 'STATE', 'ZIPCODE', 
    'PROPERTYCLASS', 'TAXDISTRICT', 'BILLINGCYCLE', 'READINGROUTE'
]
df_new = df_new.dropna(subset=required_columns)

# Remove duplicates based on LOCATIONID to simulate primary key behavior
df_new = df_new.drop_duplicates(subset='LOCATIONID')

# Ensure LOCATIONID is the first column
df_new = df_new[['LOCATIONID'] + [col for col in df_new.columns if col != 'LOCATIONID']]

# Add a trailer row with default values
trailer_row = pd.DataFrame([['TRAILER'] + [''] * (len(df_new.columns) - 1)], 
                           columns=df_new.columns)

# Append the trailer row to the DataFrame
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Save to CSV with escape character set
output_path = r"C:\Users\US97684\Downloads\documents_20250317\Outbound\STAGE_PREMISEUniq1.csv"

numeric_columns = [
    'STREETNUMBER', 'OWNERMAILSEQ', 'PROPERTYCLASS', 'TAXDISTRICT', 'BILLINGCYCLE', 'READINGROUTE',
    'SERVICEAREA', 'SERVICEFACILITY', 'PRESSUREDISTRICT', 'LATITUDE', 'LONGITUDE', 'PARCELAREATYPE',
    'PARCELAREA', 'IMPERVIOUSSQUAREFEET', 'PROPERTYUSECLASSIFICATION1', 'PROPERTYUSECLASSIFICATION2', 'AMPS', 'VOLTS', 'FLEXFIELD1', 'FLEXFIELD2',
    'PROPERTYUSECLASSIFICATION3', 'PROPERTYUSECLASSIFICATION4', 'PROPERTYUSECLASSIFICATION5'
]

def custom_quote(val, column):
    # Check if the column is in the list of numeric columns
    if column in numeric_columns:
        return val  # No quotes for numeric fields
    # Otherwise, add quotes for non-numeric fields
    return f'"{val}"' if val not in ["", None] else val

df_new = df_new.apply(lambda col: col.apply(lambda val: custom_quote(val, col.name)))

# Save to CSV with custom quoting and escape character
df_new.to_csv(output_path, index=False, quoting=csv.QUOTE_NONE, escapechar='\\')

print(f"File successfully saved to: {output_path}")
