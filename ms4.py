import pandas as pd

txt_path = 'ms4.txt'

# Read the text file and clean extra spaces by using the 'skipinitialspace' parameter
df = pd.read_csv(txt_path, delimiter='\t', skipinitialspace=True)

# Strip leading and trailing spaces from both column names and values
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Save the cleaned DataFrame to an Excel file
excel_path = 'ms4.xlsx'
df.to_excel(excel_path, index=False)

# Now read the Excel file and filter the data
df_from_excel = pd.read_excel(excel_path)

# Define the list of values to filter
filter_values = ["3/3/3 Warranty", "WARR 3/3/0 US", "WARR 3/3/3 US"]

# Strip extra spaces from column names
df_from_excel.columns = df_from_excel.columns.str.strip()

# Filter the DataFrame based on the values in the "DESCRIPTION" column
filtered_df = df_from_excel[df_from_excel['DESCRIPTION'].str.contains('|'.join(filter_values))]

# Save the filtered DataFrame to a new Excel file
filtered_excel_path = 'ms4_filtered.xlsx'
filtered_df.to_excel(filtered_excel_path, index=False)
