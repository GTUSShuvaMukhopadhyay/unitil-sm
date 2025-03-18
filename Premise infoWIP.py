import pandas as pd
import re

# Function to extract street number
def extract_street_number(address):
    address = address.strip()
    match = re.match(r"\D*(\d{1,5})(?:\s?[-\s]?\d{0,4})?", address)
    if match:
        return match.group(1)
    else:
        return ""

# Function to extract street name (excluding number and suffix)
def extract_street_name(address):
    address = address.strip()
    match = re.match(r"^\D*(?:\d{1,5}\s*[-\s]?\d{0,4})?\s*(.*)$", address)
    if match:
        return match.group(1).strip()  # Return the street name excluding the street number
    else:
        return ""  # If no match found, return an empty string

# Function to extract the last word (suffix) from the address
def extract_last_suffix(address):
    address_parts = address.strip().split()
    return address_parts[-1].upper()

# Load the data
file_path = r"Bangor test file.xlsx"
df = pd.read_excel(file_path, sheet_name='zdm_premdetails', engine='openpyxl')
df_Portion = pd.read_excel(file_path, sheet_name='TE422', engine='openpyxl')

# Load configuration file for suffix lookup
config_path = r"C:\Users\US82783\Downloads\configuration.xlsx"
sheet1 = pd.read_excel(config_path, sheet_name='Sheet1', engine='openpyxl')
sheet2 = pd.read_excel(config_path, sheet_name='Sheet2', engine='openpyxl')

# Initialize df_new as an empty DataFrame
df_new = pd.DataFrame()

# Column 1: Column B (index 1)
df_new['LOCATIONID'] = df.iloc[:, 2]

# Apply the function to extract street number from the relevant column (iloc(26))
address_data = df.iloc[:, 26].astype(str)
df_new['STREETNUMBER'] = address_data.apply(extract_street_number)

# Additional columns
df_new['STREETNAME'] = address_data.apply(extract_street_name)
df_new['DESIGNATION'] = ""  # Initially leave DESIGNATION empty
df_new['ADDITIONALDESC'] = ""  # Initially leave ADDITIONALDESC empty

df_new['TOWN'] = address_data.str.slice(0, 23)
df_new['STATE'] = "ME"
df_new['ZIPCODE'] = df.iloc[:, 27].astype(str).str.slice(0, 5)
df_new['ZIPPLUSFOUR'] = ""
df_new['OWNERCUSTOMERID'] = "TBD"
df_new['OWNERMAILSEQ'] = "1"
df_new['PROPERTYCLASS'] = "TBD"
df_new['TAXDISTRICT'] = "TBD"
df_new['BILLINGCYCLE'] = ""
df_new['READINGROUTE'] = df.iloc[:, 0].astype(str).str.slice(0, 4)
df_new['SERVICEAREA'] = ""
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

# List of valid suffixes
valid_suffixes = ["STREET", "ST", "ROAD", "RD", "LANE", "LN", "AVENUE", "AVE", "BOULEVARD", "BLVD", "CIRCLE", "CIR", "DRIVE", "DR", "COURT", "CT", "PARKWAY", "PKWY", "SQUARE", "SQR"]

# Now, update the DESIGNATION or ADDITIONALDESC based on suffix
for i in range(len(df_new)):
    address = df.iloc[i, 26]  # Get address from iloc(26) column
    
    # Extract the last suffix from the address
    suffix = extract_last_suffix(address)
    
    if suffix in valid_suffixes:
        # Check in Sheet1 for matching suffix
        matching_row_sheet1 = sheet1[sheet1.iloc[:, 0].str.upper() == suffix]  # Exact match
        
        if not matching_row_sheet1.empty:
            # If a match is found in Sheet1, set DESIGNATION
            df_new.at[i, 'DESIGNATION'] = matching_row_sheet1.iloc[0, 1]  # Assuming value is in iloc(1)
        else:
            # If no match in Sheet1, check in Sheet2
            matching_row_sheet2 = sheet2[sheet2.iloc[:, 0].str.upper() == suffix]  # Exact match
            
            if not matching_row_sheet2.empty:
                # If a match is found in Sheet2, set ADDITIONALDESC
                df_new.at[i, 'ADDITIONALDESC'] = matching_row_sheet2.iloc[0, 1]  # Assuming value is in iloc(1)

# Add a trailer row with default values
trailer_row = pd.DataFrame([['TRAILER'] + [','] * (len(df_new.columns) - 2)], 
                           columns=df_new.columns)

# Append the trailer row to the DataFrame
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Save to CSV
output_path = r"STAGE_PREMISE.csv"
df_new.to_csv(output_path, index=False)

print(f"CSV file saved at {output_path}")
