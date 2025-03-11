import pandas as pd
import os
import csv  # Import csv module for proper CSV formatting

# File path (Update accordingly)
file_path = r"\ZCAMPAIGN.xlsx"

# Read the Excel file and load the specific sheet
df = pd.read_excel(file_path, sheet_name='Sheet1', engine='openpyxl')

# Initialize df_new using relevant columns
df_new = pd.DataFrame()

# Extract relevant columns safely (Adjust column names if necessary)
df_new['CUSTOMERID'] = df.iloc[:, 1].fillna('').apply(lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x)).str.slice(0, 15)
df_new['PHONENUMBER'] = df.iloc[:, 8].fillna('').astype('str').str.slice(0, 10)

# Ensure PHONETYPE and PHONEEXT are numeric
df_new['PHONETYPE'] = 1  # Default value as numeric
df_new['PHONEEXT'] = 1  # Default value as numeric

# Add additional columns with default blank values
additional_columns = ['CONTACT', 'TITLE', 'PRIORITY', 'UPDATEDATE']
for col in additional_columns:
    df_new[col] = ""  # Use empty string instead of a space

# Add trailer row properly
trailer_row = pd.DataFrame([['TRAILER'] + [''] * (len(df_new.columns) - 1)], columns=df_new.columns)
df_new = pd.concat([df_new, trailer_row], ignore_index=True)

# Remove duplicates based on the combination of CUSTOMERID, PHONENUMBER, and PHONETYPE
df_new = df_new.drop_duplicates(subset=['CUSTOMERID', 'PHONENUMBER', 'PHONETYPE'], keep='first')

# Function to format CSV output
def custom_quote(val):
    """Returns value wrapped in quotes if not numeric, else returns as is."""
    if isinstance(val, (int, float)):  # Do not quote numeric values
        return val
    elif isinstance(val, str) and val.strip():  # If non-empty string, wrap in quotes
        return f'"{val}"'
    return val  # If empty, return as is

# Apply quoting function only to non-numeric columns
for col in df_new.columns:
    if col not in ['PHONENUMBER', 'PHONETYPE', 'PHONEEXT']:
        df_new[col] = df_new[col].apply(custom_quote)

# Define output file path
output_csv = os.path.join(os.path.dirname(file_path), 'STAGE_PHONE.csv')

# Save to CSV with proper formatting
df_new.to_csv(output_csv, index=False, header=True, quoting=csv.QUOTE_NONE, escapechar='\\')

# Confirmation message
print(f"CSV file saved at {output_csv}")
