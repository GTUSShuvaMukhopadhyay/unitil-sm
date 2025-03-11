import pandas as pd
import re
import csv

# List of common direction indicators (can be extended if needed)
directional_indicators = [
    "N", "S", "E", "W", "NE", "SE", "SW", "NW"
]

def identify_address_parts(address, street_types):
    predirection_pattern = r"(\b(?:{})\b)\s+(\w+ (?:{}))".format("|".join(directional_indicators), "|".join(street_types))
    postdirection_pattern = r"(\w+ (?:{}))\s+(\b(?:{})\b)".format("|".join(street_types), "|".join(directional_indicators))
    propername_pattern = r"(\d+[-\w]*\s+)?([A-Za-z\s\-]+(?:\s(?:{}))?)".format("|".join(street_types))

    predirection = ""
    postdirection = ""
    propername = ""

    # Check for PREDIRECTION (direction before the street name)
    predirection_match = re.search(predirection_pattern, address)
    if predirection_match:
        predirection = predirection_match.group(1)

    # Check for POSTDIRECTION (direction after the street name)
    postdirection_match = re.search(postdirection_pattern, address)
    if postdirection_match:
        postdirection = postdirection_match.group(2)

    # Check for PROPERNAME (street names)
    propername_match = re.search(propername_pattern, address)
    if propername_match:
        propername = propername_match.group(2)

    return predirection, postdirection, propername

def clean_propername(propername, directional_indicators, street_types):
    for direction in directional_indicators:
        if propername.startswith(direction + " "):
            propername = propername[len(direction) + 1:].strip()
        if propername.endswith(" " + direction):
            propername = propername[:len(propername) - len(direction) - 1].strip()

    street_name_parts = propername.split()
    if street_name_parts and street_name_parts[-1] in street_types:
        propername = " ".join(street_name_parts[:-1])  # Remove the street type

    return propername.strip()

def get_abbreviation_from_last_word(propername, directional_indicators, street_abbr_dict):
    # Check if the propername is empty before processing
    if not propername:
        return ""  # Return empty string if propername is empty

    # Split the propername into words
    street_name_parts = propername.split()

    # Ensure there are enough parts to check
    if not street_name_parts:
        return ""  # Return empty string if no street name parts

    # If the last word is a directional indicator, check the second-to-last word
    if street_name_parts[-1] in directional_indicators:
        # If the last word is a directional indicator, we look at the second-to-last word
        if len(street_name_parts) > 1 and street_name_parts[-2] in street_abbr_dict:
            return street_abbr_dict[street_name_parts[-2]]
        else:
            return ""  # Return empty string if no abbreviation found
    elif street_name_parts[-1] in street_abbr_dict:
        # If the last word is a valid street type, return the abbreviation
        return street_abbr_dict[street_name_parts[-1]]
    else:
        return ""  # Return empty string if no valid street type is found

# Read Excel file (make sure to update the file path)
file_path = r'ZMD_PremDetails_CleanData.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1', engine='openpyxl')

# Read the configuration file for street abbreviations
config_file_path = r'\configuration.xlsx'  # Update path here
config_df = pd.read_excel(config_file_path)

# Create a dictionary to map street types to their abbreviations (normalized to upper case)
street_abbr_dict = dict(zip(config_df['Description'].str.upper(), config_df['Abbreviation']))

# Extract addresses from the correct column name ('Service Address')
addresses = df['Service Address'].dropna()  # Ensure that we are selecting the correct column and skipping NaN values

# List to collect the data for CSV output
output_data = []

# Loop through the addresses and classify them
for address in addresses:
    # **Remove the city name and street number** to leave the street part only
    manipulated_address = re.sub(r'^[A-Za-z\s]+,?\s*(\d+[\s-]*\d*\s*)+', '', address).strip()

    # Use the previous logic to identify the parts of the address
    predirection, postdirection, propername = identify_address_parts(manipulated_address, street_abbr_dict.keys())

    propername = propername.upper().strip()

    # Clean PROPERNAME to exclude directional indicators and street type
    cleaned_propername = clean_propername(propername, directional_indicators, street_abbr_dict.keys())

    # Get abbreviation using the last valid street type (even if it is preceded by a directional indicator)
    abbreviation = get_abbreviation_from_last_word(propername, directional_indicators, street_abbr_dict)

    # Concatenate FULLNAME
    full_name = f"{predirection} {cleaned_propername} {abbreviation} {postdirection}".strip()

    # Create a row of data to append
    output_data.append({
        "FULLNAME": full_name,
        "PREDIRECTION": predirection,
        "PROPERNAME": cleaned_propername,
        "ABBREVIATION": abbreviation,
        "POSTDIRECTION": postdirection
    })

# Create a DataFrame from the output data
output_df = pd.DataFrame(output_data)

output_df = output_df.drop_duplicates(subset='FULLNAME')

# Create the trailer row
trailer_row = pd.DataFrame([['TRAILER'] + [''] * (len(output_df.columns) - 1)],
                           columns=output_df.columns)

# Append the trailer row to the DataFrame
output_df = pd.concat([output_df, trailer_row], ignore_index=True)

# Save the DataFrame to a CSV file with escapechar to handle special characters
def custom_quote(val):
    # Return value with quotes if it's not empty, otherwise return it as is
    if val not in ["", None]:
        return f'"{val}"'  # Add quotes around non-empty values
    return val  # Return blank values as they are

output_df = output_df.applymap(custom_quote)  # Apply the custom quoting to all cells

# Save to CSV
output_df.to_csv('STAGE_STREETS.csv', index=False, header=True, quoting=csv.QUOTE_NONE, escapechar='\\')

print("CSV file has been created and saved as 'STAGE_STREETS.csv'.")
