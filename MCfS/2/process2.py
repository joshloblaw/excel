from pandas import concat, read_excel, ExcelWriter
from fuzzywuzzy import process, fuzz
from tqdm import tqdm

ref_fid = []
exist = []

prev_fid = 0
fid = 0
is_exist = 0

md = concat(read_excel("MCfS_facilities.xlsx", sheet_name=[0], header=0, usecols=[3, 4, 5, 6, 7]), ignore_index=True)
td = concat(read_excel("refrigerant.xlsx", sheet_name=[0], header=0), ignore_index=True)

fid_db = md["Name"]
# addr_db = md["Address line 1"]

for row in tqdm(range(len(td))):
    td_fid = str(td["Facility"][row])
    # td_addr = str(td["Address"][row])

    if td_fid != prev_fid:
        md_fid = process.extractOne(td_fid, fid_db, scorer=fuzz.ratio, score_cutoff=100)
        if md_fid:
            fid = md_fid[0]
            is_exist = 1
        else:
            fid = 0
            is_exist = 0
        
    ref_fid.append(fid)
    exist.append(is_exist)
    prev_fid = td_fid

td.insert(0, "_FID", ref_fid)
td.insert(0, "_IS_EXIST", exist)

with ExcelWriter("refrigerant_output.xlsx") as writer:
    td.to_excel(writer, sheet_name="All Data")
    td.drop_duplicates("Facility").to_excel(writer, sheet_name="Facility Only")