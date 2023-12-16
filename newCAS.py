import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font

def load_excel_files(file1, file2):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    return df1, df2

def insert_rows(df1, df2, output_file="PCCS.xlsx"):


    
    # Merge dataframes based on the "Sku" and "SKU" columns
    merged_df = pd.merge(df1, df2, left_on="Sku", right_on="SKU", how="inner")
    


    # Initialize an empty list to store rows for insertion
    rows_to_insert = []

    # Iterate through the merged dataframe
    for index, row in merged_df.iterrows():
        # Check if "ContainerName" and "ContainerValue" are not NaN or floats
        if not pd.isna(row["ContainerName"]) and not pd.isna(row["ContainerValue"]):
            # For each unique combination of "ContainerName" and "ContainerValue" in df2, insert a new row
            container_names = row["ContainerName"] if isinstance(row["ContainerName"], list) else [row["ContainerName"]]
            container_values = row["ContainerValue"] if isinstance(row["ContainerValue"], list) else [row["ContainerValue"]]
            
            for container_name, container_value in zip(container_names, container_values):
                new_row = row.copy()
                new_row["ContainerName"] = container_name
                new_row["ContainerValue"] = container_value
                rows_to_insert.append(new_row)

    # Create a new dataframe with the rows to insert
    new_rows_df = pd.DataFrame(rows_to_insert)

    # Concatenate the original df1 with the new rows
    result_df = pd.concat([df1, new_rows_df], ignore_index=True)

        # List of columns to exclude
    columns_to_exclude = [
        "Sku","Name","ComponentCompletionDate", "PL", "Item_Id", "Small Series", "Big Series", "Option", "Status", 
        "CodeName", "SKU_FirstAppearanceDate", "SKU_CompletionDate", "SKU_Aging", 
        "ComponentGroup", "PhwebValue", "PhwebDescription", "ExtendedDescription", 
        "Component", "CompletionDate", "ComponentReadiness", "SKUReadiness"
    ]

    # Drop the specified columns from the merged dataframe
    result_df = result_df.drop(columns=columns_to_exclude, errors='ignore')

    result_df = result_df.rename(columns={
        "OID": "ItemId",
        "SKU": "ItemName",
        "ContainerName": "Tag",
        "ContainerValue": "ChunkValue"
    })

    # Save the result to a new Excel file with bold first row
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write the DataFrame to the Excel file
        result_df.to_excel(writer, index=False, sheet_name='Sheet1')

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
