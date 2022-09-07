import sys
import pandas as pd

# FILE NAME/PATH
MASTER_PATH = "master.xlsx"
TARGET_PATH = sys.argv[1]
OUTPUT_NAME = "OUTPUT_"  + TARGET_PATH

# SHEETS (0-indexed)
if (len(sys.argv)) == 3:
    t_sheet = int(sys.argv[2])
else:
    t_sheet = 2

MASTER_SHEETS = [4, 5, 6, 7, 8]
TARGET_SHEETS = [t_sheet]

# HEADER LOCATION (0-indexed)
MASTER_HEADER = 1
TARGET_HEADER = 0

# HEADER NAME
MASTER_NUM = "Article Number\nMARA - MATNR"
MASTER_REF = "UPC Number\nMARA - EAN11"

TARGET_NUM = "StyleNumber"
TARGET_REF = "sku"

# ----------------------------------------------#
# READING
# ----------------------------------------------#
mdict_df = pd.concat(pd.read_excel(MASTER_PATH, sheet_name=MASTER_SHEETS, header=MASTER_HEADER), ignore_index=True)
tdict_df = pd.concat(pd.read_excel(TARGET_PATH,sheet_name=TARGET_SHEETS, header=TARGET_HEADER), ignore_index=True)
pd.to_numeric(mdict_df[MASTER_REF])
pd.to_numeric(tdict_df[TARGET_REF])
mdict_df = mdict_df.sort_values(by=MASTER_REF, ignore_index=True)

# ----------------------------------------------#
# MATCHING
# ----------------------------------------------#
stylenumber = []
for i in range(len(tdict_df)):
    tupc = tdict_df[TARGET_REF][i]
    mi = mdict_df[MASTER_REF].searchsorted(value=tupc)
    man = mdict_df[MASTER_NUM][mi]
    stylenumber.append(man)

# ----------------------------------------------#
# OUTPUT
# ----------------------------------------------#
tdict_df = tdict_df.rename({TARGET_NUM:"Original"}, axis=1)
tdict_df.insert(1, TARGET_NUM, stylenumber)
tdict_df[TARGET_REF] = tdict_df[TARGET_REF].astype(str)
tdict_df2 = tdict_df.loc[tdict_df['Positive Variance'] > 0]
tdict_df3 = tdict_df.loc[tdict_df['Negative Variance'] > 0]

with pd.ExcelWriter(OUTPUT_NAME) as writer:
    tdict_df.to_excel(writer, sheet_name='Cycle Count Detail')
    tdict_df2.to_excel(writer, sheet_name='Positive Variances')
    tdict_df3.to_excel(writer, sheet_name='Negative Variances')

print("Input: " + TARGET_PATH)
print("Output: " + OUTPUT_NAME)


