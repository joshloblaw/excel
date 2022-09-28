from pandas import concat, read_excel, ExcelWriter, to_datetime, DataFrame
from fuzzywuzzy import process, fuzz
from tqdm import tqdm

md = concat(read_excel("Facilities_2.xlsx", sheet_name=[0], header=0, usecols=[3]), ignore_index=True)
td = concat(read_excel("fuel_consumption.xlsx", sheet_name=[1], header=0), ignore_index=True)

fid_db = md["Name"]
start = ['2021-01-01', '2021-02-01',  '2021-03-01', '2021-04-01', '2021-05-01', '2021-06-01', '2021-07-01', '2021-08-01', '2021-09-01', '2021-10-01', '2021-11-01', '2021-12-01']
end = ['2021-01-31', '2021-02-28', '2021-03-31', '2021-04-30', '2021-05-31', '2021-06-30', '2021-07-31', '2021-08-31', '2021-09-30', '2021-10-31', '2021-11-30', '2021-12-31']
exist = [1 for _ in range(12)]
dne = [0 for _ in range(12)]

col_exist = []
col_name = []
col_ftype = []
col_qty = []
col_qtyu = []
col_dqt = []
col_ou = []
col_facility = []
col_sdate = []
col_edate = []

for row in tqdm(range(len(td))):
    avg_qty = td["Quantity"][row] / 12
    td_fid = str(td["Facility"][row])
    md_fid = process.extractOne(td_fid, fid_db, scorer=fuzz.ratio, score_cutoff=100)
    if md_fid: 
        col_exist += exist
    else:
        col_exist += dne
    
    col_name += ["Stationary combustion" for _ in range(12)]
    col_ftype += ["Natural Gas - Residential / Commercial AB" for _ in range(12)]
    col_qty += [ avg_qty for _ in range(12) ]
    col_qtyu += ["Cubic metres" for _ in range(12)]
    col_dqt += ["Actual" for _ in range(12)]
    col_ou += ["Corporate Stores" for _ in range(12)]
    col_facility += [td_fid for _ in range(12)]
    col_sdate += start
    col_edate += end

df = DataFrame()
df.insert(0, "Consumption End Date", to_datetime(col_edate, format="%Y-%m-%d"))
df.insert(0, "Consumption Start Date", to_datetime(col_sdate, format="%Y-%m-%d"))
df.insert(0, "Facility", col_facility)
df.insert(0, "Organizational Unit", col_ou)
df.insert(0, "Data Quality Type", col_dqt)
df.insert(0, "Quantity unit", col_qtyu)
df.insert(0, "Quantity", col_qty)
df.insert(0, "Fuel Type", col_ftype)
df.insert(0, "Name", col_name)
df.insert(0, "_IS_EXIST", col_exist)

with ExcelWriter("FuelConsumptionOutput.xlsx") as writer:
    df.to_excel(writer, sheet_name="Fuel Consumption 2021")
    