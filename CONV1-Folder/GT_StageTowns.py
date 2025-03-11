import os
import pandas as pd
import csv  # Import csv module

# Define input file path
file_path = r"ZDM_PREMDETAILS.xlsx"

# Load the main dataset
df = pd.read_excel(file_path, sheet_name='Sheet1', engine='openpyxl')

# Define county mapping directly in the code
data_county = {
    "Town": ["Alton", "Argyle TWP", "Bangor", "Brewer", "Bucksport", "Chester", "Edinburg", "Frankfort", "Hampden", 
              "Hermon", "Howland", "Lincoln", "Mattamascontis TWP", "Mattawamkeag", "Old Town", "Orono", "Orrington", 
              "Prospect", "Searsport", "Veazie", "Winterport", "Woodville"],
    "Zip Code": ["04468", "04468", "04401", "04412", "04416", "04457", "04448", "04438", "04444", "04401", 
                  "04448", "04457", "04457", "04459", "04468", "04473", "04474", "04981", "04974", "04401", 
                  "04496", "04457"],
    "County": ["Penobscot", "Penobscot", "Penobscot", "Penobscot", "Hancock", "Penobscot", "Penobscot", "Waldo", "Penobscot", 
                "Penobscot", "Penobscot", "Lincoln", "Penobscot", "Penobscot", "Penobscot", "Penobscot", "Penobscot", 
                "Waldo", "Waldo", "Penobscot", "Waldo", "Penobscot"]
}

df_county = pd.DataFrame(data_county)

# Standardize Town and Zip Code formatting in df_county
df_county["Town"] = df_county["Town"].str.strip().str.lower()
df_county["Zip Code"] = df_county["Zip Code"].str.strip()

# Create an empty DataFrame for STAGE_TOWNS
df_stage_towns = pd.DataFrame()

# Extract TOWN: Everything before the first comma in "Service Address"
df_stage_towns["TOWN"] = df["Service Address"].astype(str).str.split(",").str[0].str.strip().str.lower()

# Hardcode STATE as "ME"
df_stage_towns["STATE"] = "ME"

# Assign ZIPCODE from "Zip Code" with proper formatting
df_stage_towns['ZIPCODE'] = df["Zip Code"].astype(str).str.strip().str.zfill(5)

# Merge with county reference table to get COUNTY values (preserving all rows from df_stage_towns)
df_stage_towns = df_stage_towns.merge(df_county, left_on=["TOWN", "ZIPCODE"], right_on=["Town", "Zip Code"], how="left")

df_stage_towns.drop(columns=["Town", "Zip Code"], inplace=True)  # Drop redundant columns

# Fix the chained assignment warning: use direct assignment instead of inplace=True
df_stage_towns["County"] = df_stage_towns["County"].fillna("Unknown")

# Capitalize and remove any leading/trailing whitespace from text columns
df_stage_towns["TOWN"] = df_stage_towns["TOWN"].str.upper()
df_stage_towns["STATE"] = df_stage_towns["STATE"].str.upper()
df_stage_towns["ZIPCODE"] = df_stage_towns["ZIPCODE"].str.strip()
df_stage_towns["County"] = df_stage_towns["County"].str.title()

# Apply quotes to all non-empty string values (add double quotes around text fields)
df_stage_towns = df_stage_towns.applymap(lambda x: f'"{x}"' if isinstance(x, str) and x != "" else x)

# Remove duplicates based on combination of Town, State, and Zip Code
df_stage_towns = df_stage_towns.drop_duplicates(subset=["TOWN", "STATE", "ZIPCODE"])

# Add trailer row
trailer_row = pd.DataFrame([['TRAILER'] + [''] * (len(df_stage_towns.columns) - 1)], columns=df_stage_towns.columns)
df_stage_towns = pd.concat([df_stage_towns, trailer_row], ignore_index=True)

# Dynamically define output path to avoid permission issues
output_dir = "Documents"  # Change the path to your Documents folder or another writable location
output_csv = os.path.join(output_dir, "STAGE_TOWNS.csv")

# Save to CSV with quoting enabled (using csv.QUOTE_NONE)
df_stage_towns.to_csv(output_csv, index=False, quoting=csv.QUOTE_NONE, escapechar='\\')  # quoting=csv.QUOTE_NONE disables default quoting

print(f"Data has been saved to: {output_csv}")

# Print diagnostics
total_rows = len(df_stage_towns) - 1  # Excluding trailer row
missing_county = df_stage_towns[df_stage_towns['County'] == "Unknown"]
print(f"Total rows in output: {total_rows}")
print(f"Rows with missing County: {len(missing_county)}")
