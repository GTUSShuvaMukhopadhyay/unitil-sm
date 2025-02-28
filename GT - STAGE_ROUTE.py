# GT - STAGE_ROUTE

import os
import pandas as pd

# Define the input file path
file_path = r"C:\Users\us85360\OneDrive - Grant Thornton LLP\0 - All Work\[c] 15 - Unitil\0_SM Work\BangorDataSAPSampleExtract.xlsx"

# Load the Excel sheet into a DataFrame
df = pd.read_excel(file_path, sheet_name='TE422', engine='openpyxl')

# Extract the required columns (A and B)
route_data = df.iloc[:, 0:2]
route_data.columns = ['ROUTENUMBER', 'ROUTEDESC']  # Rename the columns

# Ensure no extraneous rows (drop completely empty rows)
route_data.dropna(how='all', inplace=True)

# Define the mapping for transformation
billing_cycle_mapping = {
    "MEOTP01": "801",
    "MEOTP02": "802",
    "MEOROP01": "803",
    "MEOROP02": "804",
    "MEOROP03": "805",
    "MEBGRP01": "806",
    "MEBGRP02": "807",
    "MEBGRP03": "808",
    "MEBGRP04": "809",
    "MEBGRP05": "810",
    "MEBGRP06": "811",
    "MEBGRP07": "812",
    "MEBGRP08": "813",
    "MEBGRP09": "814",
    "MEBRWP01": "815",
    "MEBRWP02": "816",
    "MEBRWP03": "817",
    "MEBCKP01": "818",
    "MELINC01": "819",
    "METRNP01": "822"
}

# Apply the transformation
route_data['ROUTENUMBER'] = route_data['ROUTENUMBER'].replace(billing_cycle_mapping)

# Convert all non-numeric fields to strings and enclose them in double quotes
route_data = route_data.astype(str)
route_data = route_data.applymap(lambda x: f'"{x}"')

# Ensure all dates in ROUTEDESC column are in YYYY-MM-DD format (assuming ROUTEDESC might contain dates)
def enforce_date_format(value):
    try:
        return f'"{pd.to_datetime(value).strftime("%Y-%m-%d")}"'  # Convert to YYYY-MM-DD
    except Exception:
        return value  # Keep as-is if not a date

route_data['ROUTEDESC'] = route_data['ROUTEDESC'].apply(enforce_date_format)

# Replace CRLF (X'0d0a') in customer notes (ROUTEDESC column) with '~[^'
route_data['ROUTEDESC'] = route_data['ROUTEDESC'].str.replace(r'[\r\n]+', '~[^', regex=True)

# Add a 'TRAILER' row with appropriate formatting
trailer_row = pd.DataFrame([['TRAILER'] + [','] * (len(route_data.columns) - 1)], columns=route_data.columns)

route_data = pd.concat([route_data, trailer_row], ignore_index=True)

# Define the output file path
output_csv = os.path.join(os.path.dirname(file_path), 'STAGE_ROUTE.csv')

# Ensure the output filename is uppercase except for the '.csv' extension
output_csv = output_csv.upper().replace('.CSV', '.csv')

# Save the modified DataFrame to a CSV file with UTF-8 encoding
route_data.to_csv(output_csv, index=False, header=True, encoding='utf-8')

# Confirmation message
print(f"Data has been saved with transformed ROUTENUMBER values and a trailer to '{output_csv}'")
