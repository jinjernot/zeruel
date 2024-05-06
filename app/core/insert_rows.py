from app.core.get_container_value import *
from app.core.get_product import *
from openpyxl.styles import Font

import pandas as pd
import json

def insert_rows(file, output_buffer=None):
    """
    Create a new dataframe with the values found in both reports, insert the productlongname and warranty values.
    """
    df1 = pd.read_excel(file.stream, sheet_name='PRISM QUERY', engine='openpyxl')
    df2 = pd.read_excel(file.stream, sheet_name='SCS', engine='openpyxl')

    merged_df = pd.merge(df1, df2, left_on="Sku", right_on="SKU", how="inner")

    rows_to_insert = []
    processed_skus = set()

    for _, row in merged_df.iterrows():
        sku = row["Sku"]

        if sku in processed_skus:
            continue

        osinstalled_value = get_container_value(df2, sku, "osinstalled")
        processorname_value = get_container_value(df2, sku, "processorname")
        memstdes_01_value = get_container_value(df2, sku, "memstdes_01")
        hd_01des_value = get_container_value(df2, sku, "hd_01des")
        hd_02des_value = get_container_value(df2, sku, "hd_02des")
        hd_03des_value = get_container_value(df2, sku, "hd_03des")

        if (osinstalled_value is not None) and \
        (processorname_value is not None) and \
        (memstdes_01_value is not None) and \
        (hd_01des_value is not None) and \
        (hd_01des_value.strip() != ""):
            name_value = row["Name"]
            chunk_value = f"{name_value} with {osinstalled_value}, {processorname_value}, {memstdes_01_value}"

            if hd_01des_value not in ("[BLANK]", None):
                chunk_value += f", {hd_01des_value}"
            if hd_02des_value not in ("[BLANK]", None):
                chunk_value += f", {hd_02des_value}"
            if hd_03des_value not in ("[BLANK]", None):
                chunk_value += f", {hd_03des_value}"

            chunk_value += ", Low halogen"
            
            api_data = get_product(sku)

            if api_data is not None:
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

                processed_skus.add(sku)

    pccs_df = pd.DataFrame(rows_to_insert)

    ms4_filtered_df = pd.read_excel(file.stream, sheet_name='MS4', engine='openpyxl')

    with open('/opt/ais/app/python/pccs/data/warranty.json', 'r') as json_file:
        warranty_data = json.load(json_file)

    for index, row in pccs_df.iterrows():
        sku = row["Sku"]
        chunk_value = row["ChunkValue"]
        tag = row["Tag"]

        if tag == "wrntyfeatures" and any(ms4_filtered_df["SKU"].str.contains(str(sku))):
            
            matching_products = [product for warranty_list in warranty_data["warranty"] for product in warranty_list if product["Product"] == chunk_value]
            
            if matching_products:
                description = matching_products[0]["Description"]
                
                warranty_value = matching_products[0]["Warranty"]
                if any(ms4_filtered_df["DESCRIPTION                             "].str.contains(warranty_value)):
                    pccs_df.at[index, "ChunkValue"] = description

    pccs_df['ChunkStatus'] = 'F'
    pccs_df['SourceLevel'] = ''
    pccs_df['SourceCulture'] = ''
 
    if output_buffer:
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            pccs_df.to_excel(writer, index=False, sheet_name='Sheet1')

            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            for cell in worksheet["1:1"]:
                cell.font = Font(bold=True)
