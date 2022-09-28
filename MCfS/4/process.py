from pandas import concat, read_excel, ExcelWriter, to_datetime, DataFrame
from fuzzywuzzy import process, fuzz
from tqdm import tqdm

md = concat(read_excel("greenhouse.xlsx", sheet_name=[0], header=0, usecols=[3]), ignore_index=True)
td = concat(read_excel("refrigerant.xlsx", sheet_name=[0], header=0), ignore_index=True)

gas_db = md["Name"]
col_tdate = to_datetime(td["Transaction date"], format="%Y-%m-%d")

col_exist = []
col_gasref = []

for row in tqdm(range(len(td))):
    td_gas = td["Greenhouse gas"][row]
    md_gas = process.extractOne(td_gas, gas_db, scorer=fuzz.ratio, score_cutoff=100)
    if md_gas:
        col_exist.append(1)
        col_gasref.append(md_gas[0])
    else:
        col_exist.append(0)
        col_gasref.append(0)

td.insert(8, "Transaction date v2", col_tdate)
td.insert(0, "_IS_EXIST", col_exist)
td.insert(2, "_GAS_REF", col_gasref)

with ExcelWriter("output.xlsx") as writer:
    td.to_excel(writer, sheet_name="Store Refrigerant")