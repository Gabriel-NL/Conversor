import os
import pandas as pd

def convert_csv_to_xlsx(csv_file, xlsx_file):
    # Read CSV file into a pandas DataFrame
    df = pd.read_csv(csv_file)

    # Write DataFrame to Excel file
    df.to_excel(xlsx_file, index=False)

# Specify the directory containing CSV files
csv_directory = 'Input'

# Specify the directory to save XLSX files
xlsx_directory = 'Output'

# List all CSV files in the specified directory
csv_files = [file for file in os.listdir(csv_directory) if file.endswith('.csv')]

# Convert each CSV file to XLSX
for csv_file in csv_files:
    # Create the corresponding XLSX file name
    xlsx_file = os.path.join(xlsx_directory, os.path.splitext(csv_file)[0] + '.xlsx')

    # Convert CSV to XLSX
    convert_csv_to_xlsx(os.path.join(csv_directory, csv_file), xlsx_file)

print("Conversion completed.")
