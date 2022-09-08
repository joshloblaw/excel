import pandas as pd
from fuzzywuzzy import process, fuzz

# MASTER FILE VAR
MPATH = "NationalStoreList.xlsm"
MSHEET = [2] # 0-indexed
MHEAD = 11 # 0-indexed
MCOLS = [9, 10, 11, 12]
MNROWS = 1066
M_SN = "ID"
M_AD = "Address"

# TARGET FILE VAR
TPATH = "output.xlsx"
TSHEET = [1] # 0-indexed
THEAD = 0 # 0-indexed
T_SN = "Store Number"
T_AD = "Address"

# READING
mdict = pd.concat(pd.read_excel(MPATH, sheet_name=MSHEET, header=MHEAD, nrows=MNROWS, usecols=MCOLS), ignore_index=True)
tdict = pd.concat(pd.read_excel(TPATH, sheet_name=TSHEET, header=THEAD), ignore_index=True)

# MATCHING
store_numbers = []
ref_address = []
all_address = mdict[M_AD]

count, prev_t_address, address, sn = 0, 0, 0, 0

for row in range(len(tdict)):
    t_address = tdict[T_AD][row]
    if t_address == prev_t_address:
        store_numbers.append(sn)
        ref_address.append(address)
        continue

    m_address = process.extractOne(t_address, all_address, scorer=fuzz.ratio, score_cutoff=90)

    if(m_address):
        address = m_address[0]
        sn = mdict[M_SN][m_address[2]]
        count += 1
    else:
        address = -1
        sn = -1

    store_numbers.append(sn)
    ref_address.append(address)

    prev_t_address = t_address

# OUTPUT
tdict.insert(1, "Ref Address", ref_address)
tdict.insert(1, "Store Number", store_numbers)

with pd.ExcelWriter("output5.xlsx") as writer:
    tdict.to_excel(excel_writer=writer, sheet_name="2021 Electricity")

print("Completed | Count: ", count)