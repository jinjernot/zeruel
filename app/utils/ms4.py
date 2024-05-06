import pandas as pd

# Path to the text file
txt_path = 'ms4.txt'

# Read the text file into a DataFrame, using tab as delimiter and skipping initial spaces
df = pd.read_csv(txt_path, delimiter='\t', skipinitialspace=True)

# Remove extra spaces from all columns
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Path to save the Excel file
excel_path = 'ms4.xlsx'

# Save the cleaned DataFrame to an Excel file without index
df.to_excel(excel_path, index=False)

# Read the data back from the Excel file into a new DataFrame
df_from_excel = pd.read_excel(excel_path)

# List of values to filter in the "DESCRIPTION" column
filter_values = ["3/3/3 Warranty", "WARR 3/3/0 US", "WARR 3/3/3 US"]

# Strip extra spaces from column names
df_from_excel.columns = df_from_excel.columns.str.strip()

# Filter the DataFrame based on values in the "DESCRIPTION" column
filtered_df = df_from_excel[df_from_excel['DESCRIPTION'].str.contains('|'.join(filter_values))]

# Path to save the filtered Excel file
filtered_excel_path = 'ms4_filtered.xlsx'

# Save the filtered DataFrame to a new Excel file without index
filtered_df.to_excel(filtered_excel_path, index=False)
