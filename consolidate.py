import os
import openpyxl
import PySimpleGUI as sg


def consolidate(responses_folder, template_file, output_file):
    template = openpyxl.load_workbook(template_file, data_only=True)
    smelter_lookup = template["Smelter Look-up"]
    for smelter_row in smelter_lookup.rows:
        if metals.__contains__(smelter_row[0].value):
            if smelter_row[4].value is not None:
                approved_smelters.append(smelter_row[4].value)

    for file in os.listdir(responses_folder):
        xlsx_sheet = openpyxl.load_workbook(os.path.join(responses_folder, file), data_only=True)["Smelter List"]
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
    row_count = 5

    for smelter in all_smelters:
        cell_count = 1
        for cell in smelter:
            spectra_list.cell(row=row_count, column=cell_count).value = cell.value
            cell_count += 1
        row_count += 1

    template.save(output_file)


au_list = {}
sn_list = {}
w_list = {}
ta_list = {}
metals = {"Gold", "Tin", "Tungsten", "Tantalum"}
approved_smelters = []
unapproved_smelters = []

sg.theme('Reddit')
layout = [[sg.T("")],
          [sg.Text("Choose a responses folder: "), sg.Input(key="folder"), sg.FolderBrowse()],
          [sg.Text("Choose a template: "), sg.Input(key="template"), sg.FileBrowse()],
          [sg.Text("Enter output name: "), sg.Input(key="output"), sg.Text()],
          [sg.Button('Consolidate')]]
window = sg.Window('Push', layout, size=(800, 300))

while True:
    event, values = window.read()
    print(values["folder"])
    print(values["template"])
    print(values["output"])
    if event == sg.WIN_CLOSED:
        break
    elif event == 'Consolidate':
        consolidate(values["folder"], values["template"], values["output"])
        break
