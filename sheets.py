import gspread
from oauth2client.service_account import ServiceAccountCredentials
import calendar
import time
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Sabatskolas grafiks").sheet1
# Ceturkņi, lai noteiktu kurš ir aktuālais ceturksnis
ceturkshni = {
    "1" : ["Decembris", "Janvāris", "Februāris"],
    "2" : ["Marts", "Aprīlis", "Maijs"],
    "3" : ["Jūnijs", "Jūlijs", "Augusts"],
    "4" : ["Septembris", "Oktobris", "Novembris"],
}
# Mēneši un to attiecīgie index'i 
meneshi = {
    "Janvāris" : 1,
    "Februāris" : 2,
    "Marts" : 3,
    "Aprīlis" : 4, 
    "Maijs" : 5,
    "Jūnijs" : 6,
    "Jūlijs" : 7,
    "Augusts" : 8,
    "Septembris" : 9,
    "Oktobris" : 10,
    "Novembris" : 11,
    "Decembris" : 12
}
# Alfabēts, palīdzēs darbībās ar indexiem, lai zinātu kurš skaitlis nāk pēc katra
alphabet = {
    "a" : 0,
    "b" : 1,
    "c" : 2,
    "d" : 3,
    "e" : 4,
    "f" : 5,
    "g" : 6,
    "h" : 7,
    "i" : 8,
    "j" : 9,
    "k" : 10,
    "l" : 11,
    "m" : 12,
    "n" : 13,
    "o" : 14,
    "p" : 15,
    "q" : 16,
    "r" : 17,
    "s" : 18,
    "t" : 19,
    "u" : 20,
    "v" : 21,
    "w" : 22,
    "x" : 23,
    "y" : 24,
    "z" : 25
}
# Ciparu saraksts
cipari = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
# Universālala funkcija kas palīdz iegūt index'us kad ir ievadīts šūnu formāts, kā piemēram "F6:H8"
def get_RandC_indexes(string):
    burts = ""
    cipars = ""
    burts1 = ""
    cipars1 = ""
    cole = False
    for i in string:
        ir_cipars = False
        for m in cipari:
            if m == i:
                ir_cipars = True
        if i == ":":
            cole = True
        if cole == False:
            if ir_cipars == False:
                burts += i
            else:
                cipars += i
        elif cole == True and i != ":":
            if ir_cipars == False:
                burts1 += i
            else:
                cipars1 += i
    class All_Indexes:
        def __init__(self, start_column_index, end_column_index, start_row_index, end_row_index, xooo, oxoo, ooxo, ooox):
            self.start_column_index = start_column_index
            self.end_column_index = end_column_index
            self.start_row_index = start_row_index
            self.end_row_index = end_row_index
            self.xooo = xooo
            self.oxoo = oxoo
            self.ooxo = ooxo
            self.ooox = ooox
    start_column_index = 0
    end_column_index = 0
    start_row_index = int(cipars) - 1
    end_row_index = int(cipars1) - 1
    for i in alphabet:
        if i == burts.lower():
            start_column_index = alphabet[i]
        elif i == burts1.lower():
            end_column_index = alphabet[i]
    vertibas = All_Indexes(start_column_index, end_column_index, start_row_index, end_row_index, burts, cipars, burts1, cipars1)
    return vertibas
# Funkcija, kas palīdz iegūt šūnu range'u ja ir zināma sākuma šūna un tas cik collonas uz priekšu ir jāpieliek
def return_range_in_cell_format_rows(starting_cell, number):
    index = get_RandC_indexes(f"{starting_cell}:A1").xooo.lower()
    for i in alphabet:
        if i == index:
            index = alphabet[i]
    index = int(index) + int(number) - 1
    for u in alphabet:
        if alphabet[u] == index:
            index = u
    return f"{starting_cell}:{index.capitalize()}{get_RandC_indexes(f'{starting_cell}:A1').oxoo}"


# Merge'ot šūnas
def merge_cells(range_in_cell_represantation):
    a = get_RandC_indexes(range_in_cell_represantation).start_row_index
    b = get_RandC_indexes(range_in_cell_represantation).end_row_index
    c = get_RandC_indexes(range_in_cell_represantation).start_column_index
    d = get_RandC_indexes(range_in_cell_represantation).end_column_index
    body = {
    "requests": [
        {
            "mergeCells": {
                "mergeType": "MERGE_ALL",
                "range": {
                    "startRowIndex":a,
                    "endRowIndex": b + 1,
                    "startColumnIndex": c,
                    "endColumnIndex": d + 1
                }
            }
        }
    ]
    }
    return body
# Unmerge'ot šūnas
def unmerge_cells(range_in_cell_represantation):
    a = get_RandC_indexes(range_in_cell_represantation).start_row_index
    b = get_RandC_indexes(range_in_cell_represantation).end_row_index
    c = get_RandC_indexes(range_in_cell_represantation).start_column_index
    d = get_RandC_indexes(range_in_cell_represantation).end_column_index
    body = {
    "requests": [
        {
            "unmergeCells": {
                "range": {
                    "startRowIndex":a,
                    "endRowIndex": b + 1,
                    "startColumnIndex": c,
                    "endColumnIndex": d + 1
                }
            }
        }
    ]
    }
    return body

def next_or_previouse_cell_rows(cell, plus_or_minuse_one):
    letter = ""
    print(cell)
    if plus_or_minuse_one[0] == "+":
        for i in alphabet:
            if cell[0].lower() == i:
                #print(f"letter:{letter} = alphabet[i]:{alphabet[i]} alphabet[i] + 1 : {alphabet[i] + 1}    alphabet[alphabet[i] + 1] : {alphabet[alphabet[i] + 1]}")
                letter = alphabet[i] + 1
                for u in alphabet:
                    if alphabet[u] == letter:
                        letter = u
    elif plus_or_minuse_one[0] == "-":
        for i in alphabet:
            if cell[0].lower() == i:
                letter = alphabet[i] + 1
                for u in alphabet:
                    if alphabet[u] == letter:
                        letter = u
    return f"{letter}{cell[1]}".capitalize()
def colums_lenth_change(pixle_size, range):
    darbiba = {
        "requests" : [
            {
                "updateDimensionProperties" : {
                    "range" : {
                        "dimension" : "COLUMNS",
                        "startIndex" : get_RandC_indexes(range).start_column_index,
                        "endIndex" : get_RandC_indexes(range).end_column_index
                    },
                    "properties" : {
                        "pixelSize" : pixle_size
                    },
                    "fields" : "pixelSize"
                }
            }
        ]
    }
    return darbiba




def change_color(color, range):
    pass


# Mēneša klase
class Menesis:
    def __init__(self, name, all_days):
        self.cell_representation = ""
        self.all_sabaths = all_sabaths = []
        self.name = name
        self.all_days = all_days
        for i in all_days:
            if i[5] != 0:
                all_sabaths.append(i[5])
        self.number_of_sabaths = len(all_sabaths)
    def tests(self):
        print(f"\n{self.name}")
        print(f"Cik Sabati ir mēnesī {self.number_of_sabaths}")
        print(f"Kuri datumi tie ir {self.all_sabaths}")



today = f"{time.localtime().tm_mday}.{time.localtime().tm_mon}.{time.localtime().tm_year}"
year = time.localtime().tm_year
month = time.localtime().tm_mon
day = time.localtime().tm_mday
print(today)
# Iegūst aktuālo mēnesi un līdz ar to aktuālo ceturksni
aktualais_ceturksnis = ""
current_menesis = ""
for u in meneshi:
    if meneshi[u] == month:
        current_menesis = u
for j in ceturkshni:
    for m in ceturkshni[j]:
        if m == current_menesis:
            aktualais_ceturksnis = j
pirma_menesha_index = 0
otraa_menesha_index = 0
tresha_menesha_index = 0
# Iegūst datus ko ievadīt calendar bibliotēkā laidabūtu sarakstus, ko iedot mēnešu klasei
for menesis_key in meneshi:
    if menesis_key == current_menesis:
        pirma_menesha_index = meneshi[menesis_key]
        otraa_menesha_index = meneshi[menesis_key] + 1
        tresha_menesha_index = meneshi[menesis_key] + 2

menesis1 = Menesis(ceturkshni[aktualais_ceturksnis][0], calendar.monthcalendar(year, pirma_menesha_index))
menesis2 = Menesis(ceturkshni[aktualais_ceturksnis][1], calendar.monthcalendar(year, otraa_menesha_index))
menesis3 = Menesis(ceturkshni[aktualais_ceturksnis][2], calendar.monthcalendar(year, tresha_menesha_index))
menesis1.cell_representation = return_range_in_cell_format_rows("C2", menesis1.number_of_sabaths)
print(f"menesis1.cell_representation = {menesis1.cell_representation}")
menesis2.cell_representation = return_range_in_cell_format_rows(next_or_previouse_cell_rows(menesis1.cell_representation[3] + menesis1.cell_representation[4], "+1"), menesis2.number_of_sabaths)
print(f"menesis2.cell_representation = {menesis2.cell_representation}")
menesis3.cell_representation = return_range_in_cell_format_rows(next_or_previouse_cell_rows(menesis2.cell_representation[3] + menesis2.cell_representation[4], "+1"), menesis3.number_of_sabaths)
print(f"menesis3.cell_representation = {menesis3.cell_representation}")

# Iegūst mainīgo tabulu, kur tiks ievadītas vērtības
mainigaa_tabula = alphabet["c"] + menesis1.number_of_sabaths + menesis2.number_of_sabaths + menesis3.number_of_sabaths
for u in alphabet:
    if alphabet[u] == mainigaa_tabula:
        mainigaa_tabula = f"C2:{u.capitalize()}9" # Būs jāpiestrādā, kad uztaisīšu labāku funkciju, kas nosaka cik cilvēki vada Sabatskolu



#Darbibas ar tabulu
print(mainigaa_tabula)
sheet.spreadsheet.batch_update(colums_lenth_change(100, "C2:T9"))
sheet.spreadsheet.batch_update(unmerge_cells(mainigaa_tabula)) # Unmergo visu iepriekšējo tabulu
sheet.spreadsheet.batch_update(colums_lenth_change(42, mainigaa_tabula))
sheet.spreadsheet.batch_update(merge_cells(menesis1.cell_representation))
sheet.update_acell(menesis1.cell_representation, menesis1.name)
sheet.spreadsheet.batch_update(merge_cells(menesis2.cell_representation))
sheet.update_acell(menesis2.cell_representation, menesis2.name)
sheet.spreadsheet.batch_update(merge_cells(menesis3.cell_representation))
sheet.update_acell(menesis3.cell_representation, menesis3.name)

