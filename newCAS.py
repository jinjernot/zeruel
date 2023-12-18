import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

def load_excel_files(file1, file2):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    return df1, df2

def insert_rows(df1, df2, output_file="PCCS.xlsx"):

    # Merge df
    merged_df = pd.merge(df1, df2, left_on="Sku", right_on="SKU", how="inner")

    # Initialize an empty list
    rows_to_insert = []

    # Keep track of processed sku_search_values
    processed_skus = set()

    # Iterate through the merged dataframe
    for index, row in merged_df.iterrows():
        sku_search_value = row["Sku"]

        # Check if the sku_search_value has already been processed
        if sku_search_value in processed_skus:
            continue

        # Perform the search in "df2" using the selected value.
        osinstalled = df2[(df2["SKU"] == sku_search_value) & (df2["ContainerName"] == "osinstalled")]
        processorname = df2[(df2["SKU"] == sku_search_value) & (df2["ContainerName"] == "processorname")]
        memstdes_01 = df2[(df2["SKU"] == sku_search_value) & (df2["ContainerName"] == "memstdes_01")]
        hd_01des = df2[(df2["SKU"] == sku_search_value) & (df2["ContainerName"] == "hd_01des")]

        # Retrieve the values from the "ContainerValue" column of the search results.
        osinstalled_values = osinstalled["ContainerValue"].values.tolist()
        processorname_values = processorname["ContainerValue"].values.tolist()
        memstdes_01_values = memstdes_01["ContainerValue"].values.tolist()
        hd_01des_values = hd_01des["ContainerValue"].values.tolist()

        # Check if any search result is found
        if osinstalled_values and processorname_values and memstdes_01_values and hd_01des_values:
            # Use only the first matching row
            osinstalled_value = osinstalled_values[0]
            processorname_value = processorname_values[0]
            memstdes_01_value = memstdes_01_values[0]
            hd_01des_value = hd_01des_values[0]

            name_value = row["Name"]  # Get the value from the "Name" column in df1
            chunk_value = f"{name_value} with {osinstalled_value} {processorname_value} {memstdes_01_value} {hd_01des_value} Low halogen"

            # Create a new row with the extracted values
            new_row = {
                "Sku": sku_search_value,
                "Item_Id": row["Item_Id"],  # Include "Item_Id" from df1
                "ItemLevel": "Product",
                "CultureCode": "na-en",
                "DataType": "Text",
                "Tag": "prodlongname",
                "ChunkValue": chunk_value,
            }

            # Append the new row to the list
            rows_to_insert.append(new_row)

            # Add sku_search_value to the set of processed values
            processed_skus.add(sku_search_value)

    # Create a new dataframe with the rows to insert
    new_rows_df = pd.DataFrame(rows_to_insert)

    # Save the result to a new Excel file with bold first row
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write the DataFrame to the Excel file
        new_rows_df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Access the openpyxl workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Bold the first row
        for cell in worksheet["1:1"]:
            cell.font = Font(bold=True)

def main():
    file1 = "./docs/QueryReport.xlsx"
    file2 = "./docs/SCS.xlsx"

    df1, df2 = load_excel_files(file1, file2)

    # Call the function to insert rows and create the new Excel file
    insert_rows(df1, df2)

if __name__ == "__main__":
    main()
