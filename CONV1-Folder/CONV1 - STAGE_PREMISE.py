# CONV1 - STAGE_PREMISE.py
# STAGE_PREMISE.py
 
# NOTES: Update formatting
 
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
 
# Define input file path
file_path = r"MA1_Extract.xlsx"
 
# Read the Excel file and load the specific sheet
df = pd.read_excel(file_path, sheet_name='Sheet1', engine='openpyxl')
 
# Initialize df_new using relevant columns
df_new = pd.DataFrame().fillna('')
df_new['LOCATIONID'] = "TBD"
df_new['STREETNUMBER'] = "TBD"
df_new['STREETNUMBERSUFFIX'] = "TBD"
df_new['STREETNAME'] = "TBD"
df_new['DESIGNATION'] = "TBD"
df_new['ADDITIONALDESC'] = "TBD"
df_new['TOWN'] = "TBD"
df_new['STATE'] = "ME"
df_new['ZIPCODE'] = "TBD"
df_new['ZIPPLUSFOUR'] = "TBD"
df_new['OWNERCUSTOMERID'] = "TBD"
df_new['OWNERMAILSEQ'] = "TBD"
df_new['PROPERTYCLASS'] = "TBD"
df_new['TAXDISTRICT'] = "8"
df_new['BILLINGCYCLE'] = "TBD"
df_new['READINGROUTE'] = "TBD"
df_new['SERVICEAREA'] = "80"
df_new['SERVICEFACILITY'] = ""
df_new['PRESSUREDISTRICT'] = "TBD"
df_new['LATITUDE'] = "TBD"
df_new['LONGITUDE'] = "TBD"
df_new['MAPNUMBER'] = "TBD"
df_new['PARCELID'] = "TBD"
df_new['PARCELAREATYPE'] = "TBD"
df_new['PARCELAREA'] = "TBD"
df_new['IMPERVIOUSSQUAREFEET'] = "TBD"
df_new['SUBDIVISION'] = "TBD"
df_new['GISID'] = "TBD"
df_new['FOLIOSEGMENT1'] = "TBD"
df_new['FOLIOSEGMENT2'] = "TBD"
df_new['FOLIOSEGMENT3'] = "TBD"
df_new['FOLIOSEGMENT4'] = "TBD"
df_new['FOLIOSEGMENT5'] = "TBD"
df_new['PROPERTYUSECLASSIFICATION1'] = "TBD"
df_new['PROPERTYUSECLASSIFICATION2'] = "TBD"
df_new['PROPERTYUSECLASSIFICATION3'] = "TBD"
df_new['PROPERTYUSECLASSIFICATION4'] = "TBD"
df_new['PROPERTYUSECLASSIFICATION5'] = ""
df_new['UPDATEDATE'] = " "


