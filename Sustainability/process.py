import pandas as pd
from fuzzywuzzy import process, fuzz


# FILE NAME/PATH
MASTER_PATH = "Facilities.xlsx"
TARGET_PATH = "CustomReport.xlsx"
OUTPUT_NAME = "Output_" + TARGET_PATH

# SHEETS (0-indexed)
MASTER_SHEET = [0]
TARGET_SHEET = [0]

# HEADER LOCATION (0-indexed)
MASTER_HEADER = 0
TARGET_HEADER = 6

# HEADER NAME
MASTER_REF = "Address: Street 1"
MASTER_NUM = "Name"

TARGET_REF = "Address "
TARGET_NUM = "Store Number"
TARGET_ID = "EPL Identifier "

# ----------------------------------------------#
# READING
# ----------------------------------------------#
mdict_df = pd.concat(pd.read_excel(MASTER_PATH, sheet_name=MASTER_SHEET, header=MASTER_HEADER), ignore_index=True)
tdict_df = pd.concat(pd.read_excel(TARGET_PATH, sheet_name=TARGET_SHEET, header=TARGET_HEADER), ignore_index=True)

# ----------------------------------------------#
# MATCHING
# ----------------------------------------------#
store_numbers = []
master_addresses = mdict_df[MASTER_REF]
prev_epl = -1
current_epl = -1
store_number = -1

for i in range(len(tdict_df)):
    current_epl = tdict_df[TARGET_ID][i]
    if store_number != -1 and current_epl == prev_epl:
        store_numbers.append(store_number)
        continue

    print(current_epl)
    t_address = tdict_df[TARGET_REF][i]
    m_address = process.extractOne(t_address, master_addresses, scorer=fuzz.partial_ratio, score_cutoff=89)
    
    if (m_address):
        m_index = m_address[2]
        store_number = mdict_df[MASTER_NUM][m_index]
    else:
        store_number = -1

    store_numbers.append(store_number)
    prev_epl = current_epl

# ----------------------------------------------#
# OUTPUT
# ----------------------------------------------#
tdict_df.insert(0, TARGET_NUM, store_numbers)

with pd.ExcelWriter(OUTPUT_NAME) as writer:
    tdict_df.to_excel(excel_writer=writer, sheet_name="Electricity", startrow=7, startcol=1)

print("Completed")
    