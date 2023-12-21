import pandas as pd

txt_path = 'ms4.txt'
df = pd.read_csv(txt_path, delimiter='\t')

excel_path = 'ms4.xlsx'
df.to_excel(excel_path, index=False)