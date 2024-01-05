import pandas as pd
from app.core.insert_rows import *

def main():
    file1_path = "./docs/QueryReport.xlsx"
    file2_path = "./docs/SCS.xlsx"

    try:
        # Load Excel files
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)

        # Insert rows
        insert_rows(df1, df2)

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
