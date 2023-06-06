import pandas as pd
from openpyxl import load_workbook

def load_excel_files(file1, file2):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    return df1, df2

def main():
    file1 = "./docs/QueryReport.xlsx"
    file2 = "./docs/SCS.xlsx"

    df1, df2 = load_excel_files(file1, file2)

    # Load the first Excel file into a DataFrame.
    df1 = pd.read_excel(file1)

    # Remove the index from the DataFrame.
    df1.reset_index(drop=True, inplace=True)

    # Save the DataFrame to a new Excel file.
    df1.to_excel("PCCS.xlsx", index=False)

    wb = load_workbook("PCCS.xlsx")

    # Set "Sheet1" as the active sheet.
    wb.active = wb["Sheet1"]  
    sheet1 = wb.active

    # Get the content from the "Sku" column.
    sku_column = sheet1["A"]
    sku_data = [cell.value for cell in sku_column]

    # Get the content from columns D and E.
    column_d = sheet1["D"]
    column_e = sheet1["E"]
    combined_data = [str(d.value) + " " + str(e.value) for d, e in zip(column_d, column_e)]
    print (combined_data)



    # Create the "oli" sheet and paste the data into new columns.
    wb.create_sheet("oli")
    oli_sheet = wb["oli"]

    for i in range(1, len(sku_data)):
        oli_sheet.cell(row=1, column=i).value = sku_data[i]
    for i in range(1, len(combined_data)):
        oli_sheet.cell(row=2, column=i).value = combined_data[i]

    # Save the modified Excel file.
    wb.save("PCCS.xlsx")

if __name__ == "__main__":
    main()
