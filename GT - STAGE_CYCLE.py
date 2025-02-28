# CSV Staging File Checklist

CHECKLIST = [
    "✅ Filename must match the entry in Column D of the All Tables tab.",
    "✅ Filename must be in uppercase except for '.csv' extension.",
    "✅ The first record in the file must be the header row.",
    "✅ Header column names must match those in Column G of the staging table definition.",
    "✅ Ensure no extraneous rows (including blank rows) are present in the file.",
    "✅ All non-numeric fields must be enclosed in double quotes.",
    "✅ The last row in the file must be 'TRAILER' followed by commas.",
    "✅ Replace all CRLF (X'0d0a') in customer notes with ~[^",
    "✅ Ensure all dates are in 'YYYY-MM-DD' format.",
]

# Function to print the checklist
def print_checklist():
    print("CSV Staging File Validation Checklist:")
    for item in CHECKLIST:
        print(item)

# Call function to display checklist
print_checklist()






import os
import pandas as pd

# Define the input file path
file_path = r"C:\Users\us85360\OneDrive - Grant Thornton LLP\0 - All Work\[c] 15 - Unitil\0_SM Work\BangorDataSAPSampleExtract.xlsx"
# file_path = r"C:\Users\GTUSER1\Desktop\GroupA\documents_20250225\Outbound\TE420.XLSX" -  WHEN WE DO OUTBOUND IN CLIENT MACHINE



# Load the Excel sheet into a DataFrame
df = pd.read_excel(file_path, sheet_name='TE420', engine='openpyxl')




# Extract the required columns (A and B)
billing_data = df.iloc[:, 0:2]
billing_data.columns = ['BILLINGCYCLE', 'DESCRIPTION']  # Rename the columns
billing_data = billing_data.applymap(lambda x: f'"{x}"')





# Add a trailer row
trailer_row = pd.DataFrame([['TRAILER'] + [','] * (len(billing_data.columns) - 1)], columns=billing_data.columns)
billing_data = pd.concat([billing_data, trailer_row], ignore_index=True)

# Define the output file path
output_csv = os.path.join(os.path.dirname(file_path), 'STAGE_CYCLE.csv')

# Save the modified DataFrame to a CSV file
billing_data.to_csv(output_csv, index=False, header=True)

# Confirmation message
print(f"Data has been saved with renamed headers and trailer to '{output_csv}'")
