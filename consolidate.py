import os
import openpyxl

path = r"/home/lauren/Git/cmrt-consolidation/responses/"
output_file = "RMI_CMRT_6.01.xlsx"
au_list = {}
sn_list = {}
w_list = {}
ta_list = {}
metals = {"Gold", "Tin", "Tungsten", "Tantalum"}
approved_smelters = []
unapproved_smelters = []

template = openpyxl.load_workbook(output_file, data_only=True)
smelter_lookup = template["Smelter Look-up"]
for smelter_row in smelter_lookup.rows:
    if metals.__contains__(smelter_row[0].value):
        if smelter_row[4].value is not None:
            approved_smelters.append(smelter_row[4].value)

for file in os.listdir(path):
    xlsx_sheet = openpyxl.load_workbook(path+file, data_only=True)["Smelter List"]
    for r in xlsx_sheet.iter_rows():
        cid = r[0].value
        if approved_smelters.__contains__(cid):
            metal = r[1].value
            if metal == "Gold":
                au_list[cid] = r
            elif metal == "Tin":
                sn_list[cid] = r
            elif metal == "Tungsten":
                w_list[cid] = r
            elif metal == "Tantalum":
                ta_list[cid] = r
        else:
            unapproved_smelters.append(r)

all_smelters = list(au_list.values()) + list(sn_list.values()) + list(w_list.values()) + list(ta_list.values())

spectra_list = template["Smelter List"]
smelter_count = 0
row_count = 5

for smelter in all_smelters:
    cell_count = 1
    for cell in smelter:
        spectra_list.cell(row=row_count, column=cell_count).value = cell.value
        cell_count += 1
    row_count += 1

template.save("Spectra.xlsx")
