# GT - STAGE_TOWNS.py

import os
import pandas as pd

# Define input file path
file_path = r"C:\Users\us85360\OneDrive - Grant Thornton LLP\0 - All Work\[c] 15 - Unitil\0_SM Work\BangorDataSAPSampleExtract.xlsx"

# Load the main dataset
df = pd.read_excel(file_path, sheet_name='zdm_premdetails', engine='openpyxl')

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
df_stage_towns = df_stage_towns.merge(df_county, left_on=["TOWN", "ZIPCODE"], right_on=["Town", "Zip Code"], how="left") # here is the issue

df_stage_towns.drop(columns=["Town", "Zip Code"], inplace=True)  # Drop redundant columns

# Fill missing COUNTY values with "Unknown"
df_stage_towns["County"].fillna("Unknown", inplace=True)

# Ensure all text fields are quoted
df_stage_towns = df_stage_towns.applymap(lambda x: f'"{x}"' if isinstance(x, str) else x)

# Add trailer row
trailer_row = pd.DataFrame([['TRAILER'] + [','] * (len(df_stage_towns.columns) - 1)], columns=df_stage_towns.columns)
df_stage_towns = pd.concat([df_stage_towns, trailer_row], ignore_index=True)

# Dynamically define output path
output_dir = os.path.dirname(file_path)
output_csv = os.path.join(output_dir, "STAGE_TOWNS.csv")

# Save to CSV with double quotes around text fields
df_stage_towns.to_csv(output_csv, index=False, quoting=1)  # quoting=1 ensures quotes around text fields

print(f"Data has been saved to: {output_csv}")

# Print diagnostics
total_rows = len(df_stage_towns) - 1  # Excluding trailer row
missing_county = df_stage_towns[df_stage_towns['County'] == '"Unknown"']
print(f"Total rows in output: {total_rows}")
print(f"Rows with missing County: {len(missing_county)}")
