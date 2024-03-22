import pandas as pd

def merge_excel_sheets_and_save(file_path, output_csv):
    xl = pd.ExcelFile(file_path)
    sheets_dict = {}
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, skipfooter=1) 
        sheets_dict[sheet_name] = df

    merged_df = pd.concat(sheets_dict, ignore_index=True)

    merged_df.to_csv(output_csv, index=False)

    print("Merged data saved to", output_csv)


excel_file = 'ed_redemption.xlsx'
output_csv = 'ed_redemption.csv'
merge_excel_sheets_and_save(excel_file, output_csv)

excel_file = 'ed_purchase.xlsx'
output_csv = 'ed_purchase.csv'
merge_excel_sheets_and_save(excel_file, output_csv)
