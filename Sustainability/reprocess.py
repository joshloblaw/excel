import pandas as pd
from fuzzywuzzy import process, fuzz
import sys

# MASTER FILE VAR
# MPATH = "NationalStoreList.xlsm"
# MSHEET = [2]
# MHEAD = 11
# MCOLS = [9, 10, 11, 12]
# MNROWS = 1066

MPATH = "Facilities.xlsx"
MSHEET = [0]
MHEAD = 0
MCOLS = [3, 4, 5, 6]



# TARGET FILE VAR
TPATH = "output7.xlsx"
TSHEET = [0]
THEAD = 0
TCOLS = [1,2,3,4,5,6,7,8,9,10]

# HEADER
SID = "ID"
ADDR = "Address"
SN = "StoreName"
SC = "City"

# READING
master = pd.concat(pd.read_excel(MPATH, sheet_name=MSHEET, header=MHEAD, usecols=MCOLS), ignore_index=True)
target = pd.concat(pd.read_excel(TPATH, sheet_name=TSHEET, header=THEAD, usecols=TCOLS), ignore_index=True)

# MATCHING
new = []
sids = []
ref_addr = []
ref_sn = []
accs = []
cities = []
addr_book = master[ADDR]
# sn_book = master[SN]

# addr_book_lower = [ x.lower() for x in addr_book ]
# sn_book_lower = [ x.lower() for x in sn_book ]

isnew, prev_addr, prev_sn, sn, sid, addr, acc, city = 0, 0, 0, 0, 0, 0, 0, 0


for row in range (len(target)):

    if row % 1500 == 0:
        print(row)

    t_sid = target["SID"][row]
    t_addr = target["Address"][row]
    t_sn = target["StoreName"][row]

    if t_addr != prev_addr:
        if t_sid != -1 or t_sn == -1:
            sid = target["SID"][row]
            acc = target["ACCURACY"][row]
            addr = target["REF_ADDR"][row]
            sn = target["REF_SN"][row]
            city = target["REF_CITY"][row]
            isnew = 0
        else:
            m_addr = process.extractOne(t_addr.lower(), addr_book, scorer=fuzz.ratio, score_cutoff=70)
            if (m_addr):
                sn = master["StoreName"][m_addr[2]]
                acc = m_addr[1]
                addr = m_addr[0]
                sid = master["ID"][m_addr[2]]
                city = master["City"][m_addr[2]]
                isnew = 1
            else:
                sn, acc, addr, sid, city = -1, -1, -1, -1, -1
                isnew = 0

    sids.append(sid)
    ref_sn.append(sn)
    ref_addr.append(addr)
    accs.append(acc)
    cities.append(city)
    new.append(isnew)
    prev_addr = t_addr


# OUTPUT
del target["REF_CITY"]
del target["REF_SN"]
del target["REF_ADDR"]
del target["ACCURACY"]
del target["SID"]
del target["NEW"]

target.insert(0, "REF_CITY", cities)
target.insert(0, "REF_SN", ref_sn)
target.insert(0, "REF_ADDR", ref_addr)
target.insert(0, "ACCURACY", accs)
target.insert(0, "SID", sids)
target.insert(0, "NEW", new)

with pd.ExcelWriter("output8.xlsx") as writer:
    target.to_excel(excel_writer=writer, sheet_name="2021 Electricity")

print ("Completed | New Rows: ", sum(new))