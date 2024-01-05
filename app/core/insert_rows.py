from app.core.get_container_value import *
from app.core.get_product import *

from openpyxl.styles import Font
import pandas as pd

def insert_rows(df1, df2, output_file="PCCS.xlsx"):
    """Create a new dataframe with the values found in both reports, insert the productlongname and warranty values"""

    # Merge dfs
    merged_df = pd.merge(df1, df2, left_on="Sku", right_on="SKU", how="inner")

    # Initialize empty lists
    rows_to_insert = []
    processed_skus = set()

# Iterate through "merged_df"
    for _, row in merged_df.iterrows():
        sku = row["Sku"]

        # Check if the "sku" has already been processed
        if sku in processed_skus:
            continue

        # Get values from "SCS" using the get_container_value function
        osinstalled_value = get_container_value(df2, sku, "osinstalled")
        processorname_value = get_container_value(df2, sku, "processorname")
        memstdes_01_value = get_container_value(df2, sku, "memstdes_01")
        hd_01des_value = get_container_value(df2, sku, "hd_01des")

        # Check if all necessary values are found
        if osinstalled_value and processorname_value and memstdes_01_value and hd_01des_value:
            name_value = row["Name"]
            chunk_value = f"{name_value} with {osinstalled_value} {processorname_value} {memstdes_01_value} {hd_01des_value} Low halogen"

            # Get data from API using the get_product function
            api_data = get_product(sku)

            # Check if api_data is not None before accessing its elements
            if api_data is not None:
                # Create a row for prodlongname
                prodlongname_row = {
                    "Sku": sku,
                    "Item_Id": row["Item_Id"],
                    "ItemLevel": "Product",
                    "CultureCode": "na-en",
                    "DataType": "Text",
                    "Tag": "prodlongname",
                    "ChunkValue": chunk_value
                }
                rows_to_insert.append(prodlongname_row)

                if isinstance(api_data, dict):
                    chunk_value = api_data.get("name")
                else:
                    chunk_value = api_data

                warranty_row = {
                    "Sku": sku,
                    "Item_Id": row["Item_Id"],
                    "ItemLevel": "Product",
                    "CultureCode": "na-en",
                    "DataType": "Text",
                    "Tag": "warranty_features",
                    "ChunkValue": chunk_value
                }
                rows_to_insert.append(warranty_row)

                # Add sku to the set of processed values
                processed_skus.add(sku)
            else:
                print(f"Missing SKU: {sku}. Skipping...")
            # Add sku to the set of processed values
            processed_skus.add(sku)

    # Create a new dataframe with the rows to insert
    prodlongname_df = pd.DataFrame(rows_to_insert)

    # Read data from "ms4_filtered.xlsx"
    ms4_filtered_df = pd.read_excel("ms4_filtered.xlsx")

    # Load warranty data from JSON file
    with open('data/warranty.json', 'r') as json_file:
        warranty_data = json.load(json_file)

    # Iterate through the rows in prodlongname_df
    for index, row in prodlongname_df.iterrows():
        sku = row["Sku"]
        chunk_value = row["ChunkValue"]
        tag = row["Tag"]

        # Check if SKU exists in ms4_filtered_df and Tag is "warranty_features"
        if tag == "warranty_features" and any(ms4_filtered_df["SKU"].str.contains(str(sku))):
            
            # Check if ChunkValue matches any "Product" in the warranty_data
            matching_products = [product for warranty_list in warranty_data["warranty"] for product in warranty_list if product["Product"] == chunk_value]
            
            if matching_products:
                # Update prodlongname_df with the "Description" value
                description = matching_products[0]["Description"]
                
                # Additional check: Check if "Warranty" value is present in "DESCRIPTION" column of ms4_filtered_df
                warranty_value = matching_products[0]["Warranty"]
                if any(ms4_filtered_df["DESCRIPTION"].str.contains(warranty_value)):
                    prodlongname_df.at[index, "ChunkValue"] = description
 
    # Save the result to a new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write the DataFrame to the Excel file
        prodlongname_df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Access workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Bold the first row
        for cell in worksheet["1:1"]:
            cell.font = Font(bold=True)
