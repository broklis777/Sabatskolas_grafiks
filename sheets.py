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
    "1" : ["Janvāris", "Februāris", "Marts"],
    "2" : ["Aprīlis", "Maijs", "Jūnijs"],
    "3" : ["Jūlijs", "Augusts", "Septembris"],
    "4" : ["Oktobris", "Novembris", "Decembris"],
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
    class Cell_Indexes:
        def __init__ (self, start_column_index, end_column_index, start_row_index, end_row_index):
            self.start_column_index = start_column_index
            self.end_column_index = end_column_index
            self.start_row_index = start_row_index
            self.end_row_index = end_row_index
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
    if cole == False: 
        start_column_index = 0
        end_column_index = 0
        start_row_index = 0
        end_row_index = 0
        start_row_index = int(cipars) - 1 # Pilnībā nav ne jausmas kapēc
        end_row_index = start_row_index # Tā būtu + 1 jo šūnas sākums ir [ bet beigas ) bet tā kā rangeJSON funkcija jau pieskaita +1 tad nelieku šeit
        for u in alphabet:
            if u == burts.lower():
                start_column_index = alphabet[u]
                break
        end_column_index = start_column_index # Tā būtu + 1 jo šūnas sākums ir [ bet beigas ) bet tā kā rangeJSON funkcija jau pieskaita +1 tad nelieku šeit
        vertibas = Cell_Indexes(start_column_index, end_column_index, start_row_index, end_row_index)
        return vertibas
    elif cole == True:
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

# Funkcija, kas atgriež rangeu json formātā

def rangeJSON(range_in_cell_represantation):
    body = {
        "startRowIndex": get_RandC_indexes(range_in_cell_represantation).start_row_index,
        "endRowIndex": get_RandC_indexes(range_in_cell_represantation).end_row_index + 1,
        "startColumnIndex": get_RandC_indexes(range_in_cell_represantation).start_column_index,
        "endColumnIndex": get_RandC_indexes(range_in_cell_represantation).end_column_index + 1
    }
    return body

# Merge'ot šūnas
def merge_cells(range_in_cell_represantation):
    body = {
    "requests": [
        {
            "mergeCells": {
                "mergeType": "MERGE_ALL",
                "range": rangeJSON(range_in_cell_represantation)
            }
        }
    ]
    }
    return body
# Unmerge'ot šūnas
def unmerge_cells(range_in_cell_represantation):
    body = {
    "requests": [
        {
            "unmergeCells": {
                "range": rangeJSON(range_in_cell_represantation)
            }
        }
    ]
    }
    return body


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

def rows_lenth_change(pixle_size, range):
    darbiba = {
        "requests" : [
            {
                "updateDimensionProperties" : {
                    "range" : {
                        "dimension" : "ROWS",
                        "startIndex" : get_RandC_indexes(range).start_row_index,
                        "endIndex" : get_RandC_indexes(range).end_row_index
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

def mainit_krasu(range_in_cell_represantation, red, green, blue):
    # Nezin kapec sheetos indexi skaitās otrādi tapēc ir 255 - ievadītais, bet tas ar nez kapēc iedeva vairāk tapēc ir 256 - ievade
    red = 256 - red
    green = 256 - green
    blue = 256 - blue
    darbiba = {
        "requests" : [
            {
                "repeatCell" : {
                    'range': rangeJSON(range_in_cell_represantation),
                    "cell" : {
                        "userEnteredFormat" : {
                            "backgroundColor" : {
                                "red" : red,
                                "green" : green,
                                "blue" : blue
                            }
                        }
                    },
                    "fields" : "userEnteredFormat(backgroundColor)"
               }
            }
        ]
    }
    return darbiba

# Funkcija, kas addo borderus

def add_borders(range_in_cell_represantation):
    body = {
        "requests" : [
            {
                "updateBorders": {
                    "range" : rangeJSON(range_in_cell_represantation),
                    "top" : {
                        "style" : "SOLID",
                        "colorStyle" : {
                            "rgbColor" : {
                                "red" : 255,
                                "green" : 255,
                                "blue" : 255
                            }
                        }
                    },
                    "bottom" : {
                        "style" : "SOLID",
                        "colorStyle" : {
                            "rgbColor" : {
                                "red" : 255,
                                "green" : 255,
                                "blue" : 255
                            }
                        }
                    },
                    "left" : {
                        "style" : "SOLID",
                        "colorStyle" : {
                            "rgbColor" : {
                                "red" : 255,
                                "green" : 255,
                                "blue" : 255
                            }
                        }
                    },
                    "right" : {
                        "style" : "SOLID",
                        "colorStyle" : {
                            "rgbColor" : {
                                "red" : 255,
                                "green" : 255,
                                "blue" : 255
                            }
                        }
                    } 
                }
            }
        ]
    }
    return body

#Formatē šūnas tekstu
def change_font(range_in_cell_represantation:str, font_name:str, font_size:int, bold:bool, italic:bool):
    body = {
        "requests" : [
            {
                "repeatCell" : {
                    "range" : rangeJSON(range_in_cell_represantation),
                    "cell" : {
                        "userEnteredFormat" : {
                            "textFormat" : {
                                "fontFamily": font_name,
                                "fontSize": font_size,
                                "bold": bold,
                                "italic": italic,
                            }   
                        }
                    },
                    "fields" : "userEnteredFormat(textFormat)"
                }
            }
        ]
    }
    return body

#Funkcija kas maina teksta krāsu
def change_text_color(range_in_cell_represantation, red, green, blue):
    body = {
        "requests" : [
            {
                "repeatCell" : {
                    "range" : rangeJSON(range_in_cell_represantation),
                    "cell" : {
                        "userEnteredFormat" : {
                            "textFormat" : {
                                "foregroundColorStyle" : {
                                    "rgbColor" : {
                                        "red" : red,
                                        "green" : green,
                                        "blue" : blue
                                    }
                                }
                            }   
                        }
                    },
                    "fields" : "userEnteredFormat(textFormat)"
                }
            }
        ]
    }
    return body

def next_or_previouse_cell_rows(cell, plus_or_minuse_one):
    letter = ""
    if plus_or_minuse_one[0] == "+":
        for i in alphabet:
            if cell[0].lower() == i:
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

def return_an_key_from_value_of_dictionary(dictionary, value):
    for n in dictionary:
        if dictionary[n] == value:
            return n
def return_index_from_value_of_list(list, value):
    index = 0
    for i in list:
        if i == value:
            return index
        index += 1

def get_list_of_all_cells_from_range(range_in_cell_format):
    list_of_cells = []
    column = get_RandC_indexes(range_in_cell_format).xooo
    column1 = get_RandC_indexes(range_in_cell_format).ooxo
    row = int(get_RandC_indexes(range_in_cell_format).oxoo)
    row1 = int(get_RandC_indexes(range_in_cell_format).ooox)
    original_row_value = row
    starting_index = 0
    ending_index = 0 
    for i in alphabet:
        if i == column.lower():
            starting_index = alphabet[i]
        elif i == column1.lower():
            ending_index = alphabet[i]
    while starting_index < ending_index + 1:
        row = original_row_value
        while row < row1 + 1:
            list_of_cells.append(f"{return_an_key_from_value_of_dictionary(alphabet, starting_index).capitalize()}{row}")
            row += 1
        starting_index += 1
    return list_of_cells

def next_row(range_in_cell_format):
    return f"{get_RandC_indexes(range_in_cell_format).xooo}{int(get_RandC_indexes(range_in_cell_format).oxoo) + 1}:{get_RandC_indexes(range_in_cell_format).ooxo}{int(get_RandC_indexes(range_in_cell_format).ooox) + 1}"

# Vadītāju klase
class Vaditaji:
    def __init__(self, vards, epasts):
        self.vards = vards
        self.epasts = epasts
sabatskolas_vaditaji = []
vaditajs1 = Vaditaji("Luīze", "@")
vaditajs2 = Vaditaji("Māris", "@")
vaditajs3 = Vaditaji("Justs", "@")
vaditajs4 = Vaditaji("Uldis", "@")
vaditajs5 = Vaditaji("Henrijs", "@")
vaditajs6 = Vaditaji("Sindija", "@")
vaditajs7 = Vaditaji("Anastasija", "@")
sabatskolas_vaditaji.append(vaditajs1)
sabatskolas_vaditaji.append(vaditajs2)
sabatskolas_vaditaji.append(vaditajs3)
sabatskolas_vaditaji.append(vaditajs4)
sabatskolas_vaditaji.append(vaditajs5)
sabatskolas_vaditaji.append(vaditajs6)
sabatskolas_vaditaji.append(vaditajs7)
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
# Funkcija kas katrai šūnai ievada tās vērtību
def enter_values_for_range (values, cells):
    u = 0
    while u < len(values):
        print(f"cells[u] = {cells[u]}, values[u] = {values[u]}")
        sheet.update_acell(cells[u], values[u])
        u += 1
menesis1 = Menesis(ceturkshni[aktualais_ceturksnis][0], calendar.monthcalendar(year, pirma_menesha_index))
menesis2 = Menesis(ceturkshni[aktualais_ceturksnis][1], calendar.monthcalendar(year, otraa_menesha_index))
menesis3 = Menesis(ceturkshni[aktualais_ceturksnis][2], calendar.monthcalendar(year, tresha_menesha_index))
menesis1.cell_representation = return_range_in_cell_format_rows("C2", menesis1.number_of_sabaths)
menesis2.cell_representation = return_range_in_cell_format_rows(next_or_previouse_cell_rows(menesis1.cell_representation[3] + menesis1.cell_representation[4], "+1"), menesis2.number_of_sabaths)
menesis3.cell_representation = return_range_in_cell_format_rows(next_or_previouse_cell_rows(menesis2.cell_representation[3] + menesis2.cell_representation[4], "+1"), menesis3.number_of_sabaths)

# Iegūst mainīgo tabulu, kur tiks ievadītas vērtības
mainigaa_tabula = alphabet["c"] + menesis1.number_of_sabaths + menesis2.number_of_sabaths + menesis3.number_of_sabaths
for u in alphabet:
    if alphabet[u] == mainigaa_tabula:
        mainigaa_tabula = f"C2:{u.capitalize()}{3 + len(sabatskolas_vaditaji)}" # Būs jāpiestrādā, kad uztaisīšu labāku funkciju, kas nosaka cik cilvēki vada Sabatskolu

# Testa darbības ar tabulu, lai uztaisītu to tādu, kas izveidojas no jauna
sheet.clear()
# Izveido pirmo lodziņu kur rakstīts "Sabatskolas vadītāju grafiks"
sheet.spreadsheet.batch_update(rows_lenth_change(33, "A1:A2")) 
sheet.spreadsheet.batch_update(mainit_krasu("A1:E1", 255, 229, 153))
sheet.spreadsheet.batch_update(merge_cells("A1:E1"))
sheet.update_acell("A1:E1", "Sabatskolu vadītāju grafiks")
sheet.format("A1:E1", {"textFormat": {
      "foregroundColor": {
        "red": 180.0,
        "green": 180.0,
        "blue": 180.0
      },
      "fontSize": 18,
      "bold": True
    }})
sheet.spreadsheet.batch_update(add_borders("A1:E1"))
# Vadītāju lodziņš
sheet.spreadsheet.batch_update(merge_cells("A2:B3"))
sheet.spreadsheet.batch_update(add_borders("A2:B3"))
sheet.spreadsheet.batch_update(mainit_krasu("A2:B3", 162, 196, 201))
sheet.update_acell("A2:B3", "Vadītāji")
sheet.format("A2:B3", {"horizontalAlignment": "CENTER",
    "verticalAlignment": "MIDDLE",
    "textFormat": {
      "foregroundColor": {
        "red": 170.0,
        "green": 170.0,
        "blue": 170.0
      },
      "fontSize": 12,
      "bold": False
      }})

#Darbibas ar tabulu
def tabula():
    print(mainigaa_tabula)
    sheet.spreadsheet.batch_update(colums_lenth_change(100, "C2:T9")) # Uzliek kollonas normālajā izmērā
    sheet.spreadsheet.batch_update(unmerge_cells(mainigaa_tabula)) # Unmergo visu iepriekšējo tabulu
    sheet.spreadsheet.batch_update(colums_lenth_change(42, mainigaa_tabula)) # Maina visus kollonu iestatījumus lai izskatītos vairāk kā kvadrāti
    # Apvieno 1mā mēneša sūnas kā arī ievada tā vārdu, ievada nākamajā rowā Sabatu datumus, nomaina lodziņu krāsas un uzliek borderus
    sheet.spreadsheet.batch_update(merge_cells(menesis1.cell_representation))
    sheet.update_acell(menesis1.cell_representation, menesis1.name)
    enter_values_for_range(menesis1.all_sabaths, get_list_of_all_cells_from_range(next_row(menesis1.cell_representation)))
    sheet.spreadsheet.batch_update(mainit_krasu(menesis1.cell_representation, 233, 242, 250))
    sheet.spreadsheet.batch_update(mainit_krasu(next_row(menesis1.cell_representation), 233, 242, 250))
    sheet.spreadsheet.batch_update(add_borders(menesis1.cell_representation))
    sheet.spreadsheet.batch_update(add_borders(next_row(menesis1.cell_representation)))
    # Apvieno 2rā mēneša sūnas kā arī ievada tā vārdu, ievada nākamajā rowā Sabatu datumus, nomaina lodziņu krāsas un uzliek borderus
    sheet.spreadsheet.batch_update(merge_cells(menesis2.cell_representation))
    sheet.update_acell(menesis2.cell_representation, menesis2.name)
    enter_values_for_range(menesis2.all_sabaths, get_list_of_all_cells_from_range(next_row(menesis2.cell_representation)))
    sheet.spreadsheet.batch_update(mainit_krasu(menesis2.cell_representation, 233, 242, 250))
    sheet.spreadsheet.batch_update(mainit_krasu(next_row(menesis2.cell_representation), 233, 242, 250))
    sheet.spreadsheet.batch_update(add_borders(menesis2.cell_representation))
    sheet.spreadsheet.batch_update(add_borders(next_row(menesis2.cell_representation)))
    # Apvieno 3šā mēneša sūnas kā arī ievada tā vārdu, ievada nākamajā rowā Sabatu datumus, nomaina lodziņu krāsas un uzliek borderus
    sheet.spreadsheet.batch_update(merge_cells(menesis3.cell_representation))
    sheet.update_acell(menesis3.cell_representation, menesis3.name)
    enter_values_for_range(menesis3.all_sabaths, get_list_of_all_cells_from_range(next_row(menesis3.cell_representation)))
    sheet.spreadsheet.batch_update(mainit_krasu(menesis3.cell_representation, 233, 242, 250))
    sheet.spreadsheet.batch_update(mainit_krasu(next_row(menesis3.cell_representation), 233, 242, 250))
    sheet.spreadsheet.batch_update(add_borders(menesis3.cell_representation))
    sheet.spreadsheet.batch_update(add_borders(next_row(menesis3.cell_representation)))
    # Uzzīmē borderus menešu lauciņiem
    sheet.spreadsheet.batch_update(add_borders(f"{get_RandC_indexes(menesis1.cell_representation).xooo}{get_RandC_indexes(menesis1.cell_representation).oxoo}:{get_RandC_indexes(menesis1.cell_representation).ooxo}{int(get_RandC_indexes(menesis1.cell_representation).ooox) + 1 + len(sabatskolas_vaditaji)}"))
    sheet.spreadsheet.batch_update(add_borders(f"{get_RandC_indexes(menesis2.cell_representation).xooo}{get_RandC_indexes(menesis2.cell_representation).oxoo}:{get_RandC_indexes(menesis2.cell_representation).ooxo}{int(get_RandC_indexes(menesis2.cell_representation).ooox) + 1 + len(sabatskolas_vaditaji)}"))
    sheet.spreadsheet.batch_update(add_borders(f"{get_RandC_indexes(menesis3.cell_representation).xooo}{get_RandC_indexes(menesis3.cell_representation).oxoo}:{get_RandC_indexes(menesis3.cell_representation).ooxo}{int(get_RandC_indexes(menesis3.cell_representation).ooox) + 1 + len(sabatskolas_vaditaji)}"))
    # Ievada vadītājus
    time.sleep(60) #Bišķ jāuzgaida, jo savādāk pārsniedz quota limitu
    j = 0
    while j < len(sabatskolas_vaditaji):
        sheet.spreadsheet.batch_update(merge_cells(f"A{j + 4}:B{j + 4}"))
        sheet.spreadsheet.batch_update(mainit_krasu(f"A{j + 4}:B{j + 4}", 233, 242, 250))
        sheet.format(f"A{j + 4}:B{j + 4}", {"horizontalAlignment": "CENTER"})
        sheet.spreadsheet.batch_update(add_borders(f"A{j + 4}:B{j + 4}"))
        enter_values_for_range([sabatskolas_vaditaji[j].vards], [f"A{j + 4}:B{j + 4}"])
        j += 1
 
tabula()

#Mainīgā līnija
