# Import necessary modules
from app.core.get_container_value import *
from app.core.get_product import *
from openpyxl.styles import Font
import pandas as pd
import json

def insert_rows(file, output_file="PCCS.xlsx"):
    """
    Create a new dataframe with the values found in both reports, insert the productlongname and warranty values.

    Args:
        df1 (DataFrame): First DataFrame containing product information.
        df2 (DataFrame): Second DataFrame containing product information.
        output_file (str): Output filename for the resulting Excel file. Default is "PCCS.xlsx".

    Returns:
        None
    """
    df1 = pd.read_excel(file, sheet_name='PRISM QUERY', engine='openpyxl')
    df2 = pd.read_excel(file, sheet_name='SCS', engine='openpyxl')

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
        hd_02des_value = get_container_value(df2, sku, "hd_02des")
        hd_03des_value = get_container_value(df2, sku, "hd_03des")

        # Check if all necessary values are found
        if (osinstalled_value is not None) and \
        (processorname_value is not None) and \
        (memstdes_01_value is not None) and \
        (hd_01des_value is not None) and \
        (hd_01des_value.strip() != ""):
            name_value = row["Name"]
            chunk_value = f"{name_value} with {osinstalled_value}, {processorname_value}, {memstdes_01_value}"

            # Append hd_02des_value if available and not equal to "[BLANK]" or None
            if hd_01des_value not in ("[BLANK]", None):
                chunk_value += f", {hd_01des_value}"

            # Append hd_02des_value if available and not equal to "[BLANK]" or None
            if hd_02des_value not in ("[BLANK]", None):
                chunk_value += f", {hd_02des_value}"

            # Append hd_03des_value if available and not equal to "[BLANK]" or None
            if hd_03des_value not in ("[BLANK]", None):
                chunk_value += f", {hd_03des_value}"

            # Add Low halogen
            chunk_value += ", Low halogen"
            
            # Get data from API using the get_product function
            api_data = get_product(sku)

            # Check if api_data is not None before accessing its elements
            if api_data is not None:
                # Create a row for proddiff_long
                proddiff_long_row = {
                    "Sku": sku,
                    "Item_Id": row["Item_Id"],
                    "ItemLevel": "Product",
                    "CultureCode": "na-en",
                    "DataType": "Text",
                    "Tag": "proddiff_long",
                    "ChunkValue": chunk_value
                }
                rows_to_insert.append(proddiff_long_row)

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
                    "Tag": "wrntyfeatures",
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
    pccs_df = pd.DataFrame(rows_to_insert)

    # Read data from "ms4_filtered.xlsx"
    ms4_filtered_df = pd.read_excel("ms4_filtered.xlsx")

    # Load warranty data from JSON file
    with open('data/warranty.json', 'r') as json_file:
        warranty_data = json.load(json_file)

    # Iterate through the rows in pccs_df
    for index, row in pccs_df.iterrows():
        sku = row["Sku"]
        chunk_value = row["ChunkValue"]
        tag = row["Tag"]

        # Check if SKU exists in ms4_filtered_df and Tag is "wrntyfeatures"
        if tag == "wrntyfeatures" and any(ms4_filtered_df["SKU"].str.contains(str(sku))):
            
            # Check if ChunkValue matches any "Product" in the warranty_data
            matching_products = [product for warranty_list in warranty_data["warranty"] for product in warranty_list if product["Product"] == chunk_value]
            
            if matching_products:
                # Update pccs_df with the "Description" value
                description = matching_products[0]["Description"]
                
                # Additional check: Check if "Warranty" value is present in "DESCRIPTION" column of ms4_filtered_df
                warranty_value = matching_products[0]["Warranty"]
                if any(ms4_filtered_df["DESCRIPTION"].str.contains(warranty_value)):
                    pccs_df.at[index, "ChunkValue"] = description

    pccs_df['ChunkStatus'] = 'F'
    pccs_df['SourceLevel'] = ''
    pccs_df['SourceCulture'] = ''
 
    # Save the result to a new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write the DataFrame to the Excel file
        pccs_df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Access workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Bold the first row
        for cell in worksheet["1:1"]:
            cell.font = Font(bold=True)