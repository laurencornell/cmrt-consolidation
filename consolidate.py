import csv
import os
import codecs
from smelter import Smelter

path = r"/home/lauren/Git/cmrt-consolidation/responses"
au_list = {}
sn_list = {}
w_list = {}
ta_list = {}

for file in os.listdir(path):
    with codecs.open(path+'/'+file, 'r', encoding='latin-1') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        for csv_row in csv_reader:
            smelter = Smelter(csv_row[0], csv_row[1])
            if smelter.metal == "Gold":
                au_list[smelter.id_number] = smelter
            elif smelter.metal == "Tin":
                sn_list[smelter.id_number] = smelter
            elif smelter.metal == "Tungsten":
                w_list[smelter.id_number] = smelter
            elif smelter.metal == "Tantalum":
                ta_list[smelter.id_number] = smelter

all_smelters = list(au_list.values()) + list(sn_list.values()) + list(w_list.values()) + list(ta_list.values())

with open('all_smelters.csv', 'w', newline='') as final_file:
    writer = csv.writer(final_file)
    writer.writerow(['ID number', 'Metal'])
    for smelter in all_smelters:
        writer.writerow([smelter.id_number, smelter.metal])
