import pandas as pd

FILENAME="Output_CustomReport.xlsx"
OFILENAME = "Check_" + FILENAME
COLS = [1, 2, 4, 5, 7, 12]
COLNAME = "Address "

dict_df = pd.concat(pd.read_excel(FILENAME, sheet_name=[0], header=0, usecols=COLS), ignore_index=True)
dict_df = dict_df.drop_duplicates(COLNAME, keep='last')

with pd.ExcelWriter(OFILENAME) as writer:
    dict_df.to_excel(excel_writer=writer, sheet_name="Check")

print("Completed")