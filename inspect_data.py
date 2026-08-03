import pandas as pd

csv_path = r'Z:\VOC\Campaigns\2024 CHANEL\Ozone\CHANEL_DataUpload_AMSandTracegas_forTS.csv'
excel_path = r'Z:\VOC\Campaigns\2024 CHANEL\Ozone\updated_file.xlsx'

print('CSV exists:', pd.io.common.file_exists(csv_path))
print('Excel exists:', pd.io.common.file_exists(excel_path))

try:
    df_csv = pd.read_csv(csv_path, nrows=5)
    print('CSV columns:', df_csv.columns.tolist())
    print(df_csv.head(3).to_string(index=False))
except Exception as e:
    print('CSV ERROR', e)

try:
    df_xlsx = pd.read_excel(excel_path, nrows=5)
    print('Excel columns:', df_xlsx.columns.tolist())
    print(df_xlsx.head(3).to_string(index=False))
except Exception as e:
    print('Excel ERROR', e)
