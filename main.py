from config.config import ca_cert_path, client_cert_path, client_key_path, url
from openpyxl.styles import Font

import pandas as pd
import requests
import json

def load_excel_files(file1, file2):
    """Load QueryReport and SCS Accurracy files"""

    return pd.read_excel(file1), pd.read_excel(file2)

def get_container_value(df, sku, container_name):
    """Look for the container value from df2"""

    container = df[(df["SKU"] == sku) & (df["ContainerName"] == container_name)]
    values = container["ContainerValue"].values.tolist()
    return values[0] if values else None

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
                    print(f"Updated SKU: {sku}, ChunkValue with Description: {description}")
                else:
                    print(f"No match found for Warranty: {warranty_value} in SKU: {sku}, ChunkValue: {chunk_value}")
            else:
                print(f"No match found for SKU: {sku}, ChunkValue: {chunk_value}")


    

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




def get_product(sku):
    """Get products from API"""
    try:
        ca_cert = ca_cert_path
        client_cert = (client_cert_path, client_key_path)

        json_data = {
            "sku": [sku],
            "countryCode": "US",
            "languageCode": "EN",
            "layoutName": "PDPCOMBO",
            "requestor": "HERMESQA-PRO",
            "reqContent": ["hierarchy"]
        }

        response = requests.post(
            url,
            cert=client_cert,
            verify=ca_cert,
            headers={'Content-Type': 'application/json'},
            data=json.dumps(json_data)
        )

        if response.status_code == 200:
            # Save the response to a JSON file
            #json_filename = f"response_{sku}.json"
            #with open(json_filename, 'w') as json_file:
            #    json.dump(response.json(), json_file)

            response_data = response.json()
            
            marketing_category = response_data["products"][sku]["productHierarchy"]["marketingCategory"].get("name")

            if marketing_category == "Workstations":
                marketing_sub_category = response_data["products"][sku]["productHierarchy"]["marketingSubCategory"].get("name")
                
                if marketing_sub_category == "HP Mobile Workstation":
                    # Use marketing_sub_category value instead of marketing_category
                    api_data = marketing_sub_category
                else:
                    api_data = marketing_category
            else:
                api_data = marketing_category

            return api_data


        else:
            print(f"Request failed with status code {response.status_code}")
            print(f"Response content: {response.text}")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    file1 = "./docs/QueryReport.xlsx"
    file2 = "./docs/SCS.xlsx"

    df1, df2 = load_excel_files(file1, file2)

    insert_rows(df1, df2)

if __name__ == "__main__":
    main()
