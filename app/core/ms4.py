import pandas as pd

txt_path = 'ms4.txt'

# Read the text file
df = pd.read_csv(txt_path, delimiter='\t', skipinitialspace=True)

# Remove extra spaces
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Save to a new Excel file
excel_path = 'ms4.xlsx'
df.to_excel(excel_path, index=False)

# read the Excel file
df_from_excel = pd.read_excel(excel_path)

# Define the list of values to filter
filter_values = ["3/3/3 Warranty", "WARR 3/3/0 US", "WARR 3/3/3 US"]

# Strip extra spaces from column names
df_from_excel.columns = df_from_excel.columns.str.strip()

# Filter based on the values in the "DESCRIPTION" column
filtered_df = df_from_excel[df_from_excel['DESCRIPTION'].str.contains('|'.join(filter_values))]

# Save to a new Excel file
filtered_excel_path = 'ms4_filtered.xlsx'
filtered_df.to_excel(filtered_excel_path, index=False)
