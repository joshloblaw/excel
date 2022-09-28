from pandas import concat, read_excel, ExcelWriter
from fuzzywuzzy import process, fuzz
from tqdm import tqdm

# md_id = "Name"
# td_id = "Facility"
# md_addr = "Address line 1"
# td_addr = "Address"
# md_name = "Address line 2"
# td_name = "Store Name"
# d_city = "City"

ref_fid = []
ref_addr = []
suggestions = []

prev_fid = 0
fid = 0
addr = 0
is_sug = 0

md = concat(read_excel("MCfS_facilities.xlsx", sheet_name=[0], header=0, usecols=[3, 4, 5, 6, 7]), ignore_index=True)
td = concat(read_excel("electricity.xlsx", sheet_name=[0], header=0, usecols=[1, 2, 3, 4, 5, 6, 7, 9]), ignore_index=True)

fid_db = md["Name"]
addr_db = md["Address line 1"]

for row in tqdm(range(len(td))):
    td_fid = str(td["Facility"][row])
    td_addr = str(td["Address"][row])

    if td_fid != prev_fid:
        md_fid = process.extractOne(td_fid, fid_db, scorer=fuzz.ratio, score_cutoff=100)
        if md_fid:
            fid = md_fid[0]
            addr = md["Address line 1"][md_fid[2]]
            is_sug = 0
        else:
            md_fid = process.extractOne(td_addr, addr_db, scorer=fuzz.ratio, score_cutoff=90)
            if md_fid:
                fid = md["Name"][md_fid[2]]
                addr = md_fid[0]
                is_sug = 1
            else:
                fid = 0
                addr = 0
                is_sug = 0
        
    ref_fid.append(fid)
    ref_addr.append(addr)

    suggestions.append(is_sug)
    prev_fid = td_fid

td.insert(0, "_REF_ADDR", ref_addr)
td.insert(0, "_FID", ref_fid)
td.insert(0, "_SUGGESTION", suggestions)

with ExcelWriter("output_with_suggestion.xlsx") as writer:
    td.to_excel(writer, sheet_name="All Data")
    td.drop_duplicates("Facility").to_excel(writer, sheet_name="Facility Only")