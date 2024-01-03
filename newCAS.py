import pandas as pd
from openpyxl.styles import Font
from config import ca_cert_path, client_cert_path, client_key_path, url
import requests
import json


def load_excel_files(file1, file2):
    return pd.read_excel(file1), pd.read_excel(file2)

def get_container_value(df, sku, container_name):
    container = df[(df["SKU"] == sku) & (df["ContainerName"] == container_name)]
    values = container["ContainerValue"].values.tolist()
    return values[0] if values else None

def insert_rows(df1, df2, output_file="PCCS.xlsx"):
    # Merge dfs
    merged_df = pd.merge(df1, df2, left_on="Sku", right_on="SKU", how="inner")

    # Initialize empty list
    rows_to_insert = []

    processed_skus = set()

# Iterate through "merged_df"
    for _, row in merged_df.iterrows():
        sku = row["Sku"]

        # Check if the "sku" has already been processed
        if sku in processed_skus:
            continue

        # Get values from "df2" using the helper function
        osinstalled_value = get_container_value(df2, sku, "osinstalled")
        processorname_value = get_container_value(df2, sku, "processorname")
        memstdes_01_value = get_container_value(df2, sku, "memstdes_01")
        hd_01des_value = get_container_value(df2, sku, "hd_01des")

        # Check if all necessary values are found
        if osinstalled_value and processorname_value and memstdes_01_value and hd_01des_value:
            name_value = row["Name"]
            chunk_value = f"{name_value} with {osinstalled_value} {processorname_value} {memstdes_01_value} {hd_01des_value} Low halogen"

            # Get big_series_data using the helper function
            big_series_data = get_product(sku)

            # Check if big_series_data is not None before accessing its elements
            if big_series_data is not None:
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

                # Create a row for warranty_features
                warranty_row = {
                    "Sku": sku,
                    "Item_Id": row["Item_Id"],
                    "ItemLevel": "Product",
                    "CultureCode": "na-en",
                    "DataType": "Text",
                    "Tag": "warranty_features",
                    "ChunkValue": big_series_data.get("name")
                }
                rows_to_insert.append(warranty_row)

                # Add sku to the set of processed values
                processed_skus.add(sku)
            else:
                print(f"big_series_data is None for SKU: {sku}. Skipping...")
            # Add sku to the set of processed values
            processed_skus.add(sku)

    # Create a new dataframe with the rows to insert
    prodlongname_df = pd.DataFrame(rows_to_insert)

    # Save the result to a new Excel file with bold first row
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write the DataFrame to the Excel file
        prodlongname_df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Access the openpyxl workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Bold the first row
        for cell in worksheet["1:1"]:
            cell.font = Font(bold=True)

def get_product(sku):
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
            #print("Request successful!")

            # Save the response to a JSON file
            json_filename = f"response_{sku}.json"
            with open(json_filename, 'w') as json_file:
                json.dump(response.json(), json_file)

            response = response.json()
            big_series_data = response["products"][sku]["productHierarchy"]["marketingCategory"]

            return big_series_data
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
