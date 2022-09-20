import pandas as pd
from fuzzywuzzy import process, fuzz
import sys

# MASTER FILE VAR
MPATH = "NationalStoreList.xlsm"
MSHEET = [2]
MHEAD = 11
MCOLS = [9, 10, 11, 12]
MNROWS = 1066

# TARGET FILE VAR
TPATH = "CustomReport2021_v2.xlsx"
TSHEET = [0]
THEAD = 0
TCOLS = [0, 1, 2, 3, 17, 24]

# HEADER
ID = "ID"
ADDR = "Address"
SN = "StoreName"
SC = "City"

# READING
master = pd.concat(pd.read_excel(MPATH, sheet_name=MSHEET, header=MHEAD, nrows=MNROWS, usecols=MCOLS), ignore_index=True)
target = pd.concat(pd.read_excel(TPATH, sheet_name=TSHEET, header=THEAD, usecols=TCOLS), ignore_index=True)

# MATCHING
sids = []
ref_addr = []
ref_sn = []
accs = []
cities = []
addr_book = master[ADDR]
sn_book = master[SN]
sn_book_lower = [ x.lower() for x in sn_book ]

count, prev_sn, sn, sid, addr, acc, city = 0, 0, 0, 0, 0, 0, 0

for row in range (len(target)):

    if row % 1500 == 0:
        print(row)

    t_sn = target[SN][row]

    if t_sn == -1:
        sids.append(-1)
        ref_sn.append(-1)
        ref_addr.append(-1)
        accs.append(-1)
        cities.append(-1)
        continue
    else:
        try:
            t_sn.lower()
        except AttributeError:
            print(t_sn, row)
            sys.exit(1)


    if t_sn == prev_sn:

        if (sid != -1):
            count += 1

        sids.append(sid)
        ref_sn.append(sn)
        ref_addr.append(addr)
        accs.append(acc)
        cities.append(city)

        continue

    m_sn = process.extractOne(t_sn, sn_book, scorer=fuzz.ratio, score_cutoff=100)

    if (m_sn):
        sn = m_sn[0]
        acc = m_sn[1]
        addr = master[ADDR][m_sn[2]]
        sid = master[ID][m_sn[2]]
        city = master[SC][m_sn[2]]
        count += 1
    else:
        sn = -1
        acc = -1
        addr = -1
        sid = -1
        city = -1

    sids.append(sid)
    ref_sn.append(sn)
    ref_addr.append(addr)
    accs.append(acc)
    cities.append(city)

    prev_sn = t_sn

# OUTPUT
target.insert(0, "REF_CITY", cities)
target.insert(0, "REF_SN", ref_sn)
target.insert(0, "REF_ADDR", ref_addr)
target.insert(0, "ACCURACY", accs)
target.insert(0, "SID", sids)

with pd.ExcelWriter("output.xlsx") as writer:
    target.to_excel(excel_writer=writer, sheet_name="2021 Electricity")

print ("Completed | Count: ", count)