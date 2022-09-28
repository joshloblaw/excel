from re import T
import re
import pandas as pd
from fuzzywuzzy import fuzz

# nsl = pd.concat(pd.read_excel("NationalStoreList.xlsm", sheet_name=[2], usecols=[9, 10, 11, 12], nrows=1066, header=11), ignore_index=True)
df = pd.concat(pd.read_excel("stores.xlsx", sheet_name=[0], header=0), ignore_index=True)

x = []

for row in range(len(df)):
    try:
        addr = df["Address"][row].lower()
        ref_addr = df["REF_ADDR"][row].lower()
    except AttributeError:
        x.append(-1)
        continue

    if fuzz.ratio(addr, ref_addr) > 94:
        x.append(1)
    elif fuzz.partial_ratio(addr, ref_addr) > 94:
        x.append(2)
    else:
        x.append(0)
    
df.insert(0, "x", x)

with pd.ExcelWriter("out.xlsx") as writer:
    df.to_excel(excel_writer=writer, sheet_name="c")

