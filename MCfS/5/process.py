from pandas import concat, read_excel, ExcelWriter, to_datetime
from tqdm import tqdm
import datetime as dt

md = concat(read_excel("emission2.xlsx", sheet_name=[
            0], header=0), ignore_index=True)
td = concat(read_excel("refrigerant2021.xlsx",
            sheet_name=[0], header=0), ignore_index=True)

md['Transaction date'] = to_datetime(md['Transaction date'])
md = md[md['Transaction date'].dt.year == 2021]
md = md[md['Organizational unit'] == 'Corporate Stores']

col_dup = []
col_fid = []
col_ghg = []
col_qty = []
col_tdate = []


db = md['Name']

for td_row in tqdm(range(len(td))):
    score, fid, ghg, qty, tdate = 0, 0, 0, 0, 0
    td_fid = str(td['Name'][td_row])
    td_ghg = str(td['Greenhouse gas'][td_row])
    td_qty = float(td['Quantity'][td_row])
    td_td = to_datetime(td['Transaction date'][td_row])

    for md_row in range(len(md)):
        md_fid = str(md['Name'][md_row])
        md_ghg = str(md['Greenhouse gas'][md_row])
        md_qty = float(md['Quantity'][md_row])
        md_td = to_datetime(md['Transaction date'][md_row])

        if td_fid == md_fid:
            score = 1
            fid = md_fid
            if (td_ghg == md_ghg) and (td_qty == md_qty) and (td_td == md_td):
                score += 1
                ghg = md_ghg
                qty = md_qty
                tdate = md_td
                break

    col_dup.append(score)
    col_fid.append(fid)
    col_ghg.append(ghg)
    col_qty.append(qty)
    col_tdate.append(tdate)


td.insert(0, "_TD_REF", col_tdate)
td.insert(0, "_QTY_REF", col_qty)
td.insert(0, "_GHG_REF", col_ghg)
td.insert(0, "_FID_REF", col_fid)
td.insert(0, "_IS_DUP", col_dup)

with ExcelWriter("output2.xlsx") as writer:
    td.to_excel(writer, sheet_name="Duplicates")
