from turtle import Turtle
import pandas as pd
from fuzzywuzzy import process, fuzz

MASTER_PATH = "Facilities.xlsx"
TARGET_PATH = "Debug3_CustomReport.xlsx"
OUTPUT_NAME = "Debug4_CustomReport.xlsx"

TARGET_SHEET = [1]
MASTER_SHEET = [0]

MASTER_REF = "Address: Street 1"
MASTER_NUM = "Name"

TARGET_REF = "Address "
TARGET_NUM = "Store Number"

tdict = pd.concat(pd.read_excel(TARGET_PATH, sheet_name=TARGET_SHEET, header=0), ignore_index=True)
mdict = pd.concat(pd.read_excel(MASTER_PATH, sheet_name=MASTER_SHEET, header=0), ignore_index=True)

store_numbers = []
matched_addresses = []
accuracies = []
master_addresses = mdict[MASTER_REF]

for i in range(len(tdict)):
    t_address = tdict[TARGET_REF][i]
    m_address = process.extractOne(t_address, master_addresses, scorer=fuzz.partial_ratio)
    extracted_address = m_address[0]
    extracted_accuracy = m_address[1]
    extracted_index = m_address[2]
    store_number = mdict[MASTER_NUM][extracted_index]
    store_numbers.append(store_number)
    matched_addresses.append(extracted_address)
    accuracies.append(extracted_accuracy)

tdict.insert(0, "Accuracy2", accuracies)
tdict.insert(0, "MatchedAddress2", matched_addresses)
tdict.insert(0, "STORENUMBER2", store_numbers)

with pd.ExcelWriter(OUTPUT_NAME) as writer:
    tdict.to_excel(excel_writer=writer, sheet_name="Electricity")