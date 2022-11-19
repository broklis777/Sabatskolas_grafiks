from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import calendar
import time
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Sabatskolas grafiks").sheet1
pagajusa_ceturksna_dati = {
    "gads" : 2022,
    "ceturksnis" : 4,
    "seciba" : ["Luīze", "Uldis", "Justs", "Māris", "Sindija", "Henrijs", "Luīze", "Māris", "Justs", "Uldis", "Henrijs", "Sindija", "Luīze", "Māris"]
}
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
        def __init__ (self, start_column_index, end_column_index, start_row_index, end_row_index, xo, ox):
            self.start_column_index = start_column_index
            self.end_column_index = end_column_index
            self.start_row_index = start_row_index
            self.end_row_index = end_row_index
            self.xo = xo
            self.ox = ox
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
        vertibas = Cell_Indexes(start_column_index, end_column_index, start_row_index, end_row_index, burts, cipars)
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

def add_borders(range_in_cell_represantation, add_or_remove):
    if add_or_remove == "+":
        a = "SOLID"
    elif add_or_remove == "-":
        a = "NONE"
    body = {
        "requests" : [
            {
                "updateBorders": {
                    "range" : rangeJSON(range_in_cell_represantation),
                    "top" : {
                        "style" : a,
                        "colorStyle" : {
                            "rgbColor" : {
                                "red" : 255,
                                "green" : 255,
                                "blue" : 255
                            }
                        }
                    },
                    "bottom" : {
                        "style" : a,
                        "colorStyle" : {
                            "rgbColor" : {
                                "red" : 255,
                                "green" : 255,
                                "blue" : 255
                            }
                        }
                    },
                    "left" : {
                        "style" : a,
                        "colorStyle" : {
                            "rgbColor" : {
                                "red" : 255,
                                "green" : 255,
                                "blue" : 255
                            }
                        }
                    },
                    "right" : {
                        "style" : a,
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
        #print(f"list is {list}")
        #print(f"value is {value}")
        #if i == value:
            #print(f"i == value     {i} == {value}     and index is {index}")
        #if i == 26 and value == 26:
            #print("nu vo")
        if int(i) == int(value) or str(i) == str(value):
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
        #elif i == column1.lower():
            #ending_index = alphabet[i]
    for u in alphabet:
        if u == column1.lower():
            ending_index = alphabet[u]
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
        self.cell_representation = "T10"
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
y = 0
while y < len(sabatskolas_vaditaji):
    sabatskolas_vaditaji[y].cell_representation = f"A{y + 4}:B{y + 4}"
    y += 1
# Mēneša klase
class Menesis:
    def __init__(self, name, all_days):
        self.cell_representation = ""
        self.all_sabaths = all_sabaths = []
        self.name = name
        self.all_days = all_days
        self.index = 0
        for i in all_days:
            if i[5] != 0:
                all_sabaths.append(i[5])
        self.number_of_sabaths = len(all_sabaths)


# Funkcija kas katrai šūnai ievada tās vērtību
def enter_values_for_range (values, cells):
    u = 0
    while u < len(values):
        #print(f"cells[u] = {cells[u]}, values[u] = {values[u]}")
        sheet.update_acell(cells[u], values[u])
        u += 1


today = f"{time.localtime().tm_mday}.{time.localtime().tm_mon}.{time.localtime().tm_year}"
#year = time.localtime().tm_year
#month = time.localtime().tm_mon
#day = time.localtime().tm_mday

def galvenas_operacijas():
    # Iegūst aktuālo mēnesi un līdz ar to aktuālo ceturksni
    global aktualais_ceturksnis
    global current_menesis
    aktualais_ceturksnis = ""
    current_menesis = ""

    for u in meneshi:
        print(f"meneshi[u] {meneshi[u]} == month {month}")
        if meneshi[u] == month:
            current_menesis = u
    for j in ceturkshni:
        for m in ceturkshni[j]:
            if m == current_menesis:
                aktualais_ceturksnis = j
    print(f"Current menesis ir {current_menesis}")
    global pirma_menesha_index
    global otraa_menesha_index
    global tresha_menesha_index
    pirma_menesha_index = 0
    otraa_menesha_index = 0
    tresha_menesha_index = 0
    # Iegūst datus ko ievadīt calendar bibliotēkā laidabūtu sarakstus, ko iedot mēnešu klasei
    current_menesha_index = 0
    for menesis_key in meneshi:
        if menesis_key == current_menesis:
            current_menesha_index = meneshi[menesis_key]
            print(f"{current_menesha_index} = {meneshi[menesis_key]}")
    if current_menesis == ceturkshni[aktualais_ceturksnis][0]:
        print("izpildijās")
        pirma_menesha_index = current_menesha_index
        otraa_menesha_index = pirma_menesha_index + 1
        tresha_menesha_index = pirma_menesha_index +2
    elif current_menesis == ceturkshni[aktualais_ceturksnis][1]:
        pirma_menesha_index = current_menesha_index - 1
        otraa_menesha_index = current_menesha_index
        tresha_menesha_index = current_menesha_index + 1
    else:
        pirma_menesha_index = current_menesha_index - 2
        otraa_menesha_index = current_menesha_index - 1
        tresha_menesha_index = current_menesha_index
    global menesis1
    global menesis2
    global menesis3
    print(f"Pirma meneshaaaaa index ir {pirma_menesha_index}")
    menesis1 = Menesis(ceturkshni[aktualais_ceturksnis][0], calendar.monthcalendar(year, pirma_menesha_index))
    menesis2 = Menesis(ceturkshni[aktualais_ceturksnis][1], calendar.monthcalendar(year, otraa_menesha_index))
    menesis3 = Menesis(ceturkshni[aktualais_ceturksnis][2], calendar.monthcalendar(year, tresha_menesha_index))
    menesis1.cell_representation = return_range_in_cell_format_rows("C2", menesis1.number_of_sabaths)
    menesis2.cell_representation = return_range_in_cell_format_rows(next_or_previouse_cell_rows(menesis1.cell_representation[3] + menesis1.cell_representation[4], "+1"), menesis2.number_of_sabaths)
    menesis3.cell_representation = return_range_in_cell_format_rows(next_or_previouse_cell_rows(menesis2.cell_representation[3] + menesis2.cell_representation[4], "+1"), menesis3.number_of_sabaths)
    menesis1.index = pirma_menesha_index
    menesis2.index = otraa_menesha_index
    menesis3.index = tresha_menesha_index

    # Iegūst mainīgo tabulu, kur tiks ievadītas vērtības
    global mainigaa_tabula
    mainigaa_tabula = alphabet["c"] + menesis1.number_of_sabaths + menesis2.number_of_sabaths + menesis3.number_of_sabaths
    for u in alphabet:
        if alphabet[u] == mainigaa_tabula:
            mainigaa_tabula = f"C2:{u.capitalize()}{3 + len(sabatskolas_vaditaji)}" # Būs jāpiestrādā, kad uztaisīšu labāku funkciju, kas nosaka cik cilvēki vada Sabatskolu


# Nodzēst pilnībā visu
def clear_all():
    sheet.spreadsheet.batch_update(merge_cells("A1:W30"))
    sheet.spreadsheet.batch_update(mainit_krasu("A1:W30", 255, 255, 255))
    sheet.format("A1:W30", {
        "horizontalAlignment": "CENTER",
        "textFormat": {
        "foregroundColor": {
            "red": 0.0,
            "green": 0.0,
            "blue": 0.0
        },
        "fontSize": 10,
        "bold": False
    }})
    sheet.spreadsheet.batch_update(add_borders("A1:W30", "-"))
    sheet.spreadsheet.batch_update(unmerge_cells("A1:W30"))
    sheet.clear()

#Darbibas ar tabulu
def tabula():
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
    sheet.spreadsheet.batch_update(add_borders("A1:E1", "+"))
    # Vadītāju lodziņš
    sheet.spreadsheet.batch_update(merge_cells("A2:B3"))
    sheet.spreadsheet.batch_update(add_borders("A2:B3", "+"))
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
    sheet.spreadsheet.batch_update(colums_lenth_change(100, "C2:T9")) # Uzliek kollonas normālajā izmērā
    sheet.spreadsheet.batch_update(unmerge_cells(mainigaa_tabula)) # Unmergo visu iepriekšējo tabulu
    sheet.spreadsheet.batch_update(colums_lenth_change(42, mainigaa_tabula)) # Maina visus kollonu iestatījumus lai izskatītos vairāk kā kvadrāti
    # Apvieno 1mā mēneša sūnas kā arī ievada tā vārdu, ievada nākamajā rowā Sabatu datumus, nomaina lodziņu krāsas un uzliek borderus
    sheet.spreadsheet.batch_update(merge_cells(menesis1.cell_representation))
    sheet.update_acell(menesis1.cell_representation, menesis1.name)
    enter_values_for_range(menesis1.all_sabaths, get_list_of_all_cells_from_range(next_row(menesis1.cell_representation)))
    sheet.spreadsheet.batch_update(mainit_krasu(menesis1.cell_representation, 233, 242, 250))
    sheet.spreadsheet.batch_update(mainit_krasu(next_row(menesis1.cell_representation), 233, 242, 250))
    sheet.spreadsheet.batch_update(add_borders(menesis1.cell_representation, "+"))
    sheet.spreadsheet.batch_update(add_borders(next_row(menesis1.cell_representation), "+"))
    # Apvieno 2rā mēneša sūnas kā arī ievada tā vārdu, ievada nākamajā rowā Sabatu datumus, nomaina lodziņu krāsas un uzliek borderus
    sheet.spreadsheet.batch_update(merge_cells(menesis2.cell_representation))
    sheet.update_acell(menesis2.cell_representation, menesis2.name)
    enter_values_for_range(menesis2.all_sabaths, get_list_of_all_cells_from_range(next_row(menesis2.cell_representation)))
    sheet.spreadsheet.batch_update(mainit_krasu(menesis2.cell_representation, 233, 242, 250))
    sheet.spreadsheet.batch_update(mainit_krasu(next_row(menesis2.cell_representation), 233, 242, 250))
    sheet.spreadsheet.batch_update(add_borders(menesis2.cell_representation, "+"))
    sheet.spreadsheet.batch_update(add_borders(next_row(menesis2.cell_representation), "+"))
    # Apvieno 3šā mēneša sūnas kā arī ievada tā vārdu, ievada nākamajā rowā Sabatu datumus, nomaina lodziņu krāsas un uzliek borderus
    sheet.spreadsheet.batch_update(merge_cells(menesis3.cell_representation))
    sheet.update_acell(menesis3.cell_representation, menesis3.name)
    enter_values_for_range(menesis3.all_sabaths, get_list_of_all_cells_from_range(next_row(menesis3.cell_representation)))
    sheet.spreadsheet.batch_update(mainit_krasu(menesis3.cell_representation, 233, 242, 250))
    sheet.spreadsheet.batch_update(mainit_krasu(next_row(menesis3.cell_representation), 233, 242, 250))
    sheet.spreadsheet.batch_update(add_borders(menesis3.cell_representation, "+"))
    sheet.spreadsheet.batch_update(add_borders(next_row(menesis3.cell_representation), "+"))
    # Uzzīmē borderus menešu lauciņiem
    sheet.spreadsheet.batch_update(add_borders(f"{get_RandC_indexes(menesis1.cell_representation).xooo}{get_RandC_indexes(menesis1.cell_representation).oxoo}:{get_RandC_indexes(menesis1.cell_representation).ooxo}{int(get_RandC_indexes(menesis1.cell_representation).ooox) + 1 + len(sabatskolas_vaditaji)}", "+"))
    sheet.spreadsheet.batch_update(add_borders(f"{get_RandC_indexes(menesis2.cell_representation).xooo}{get_RandC_indexes(menesis2.cell_representation).oxoo}:{get_RandC_indexes(menesis2.cell_representation).ooxo}{int(get_RandC_indexes(menesis2.cell_representation).ooox) + 1 + len(sabatskolas_vaditaji)}", "+"))
    sheet.spreadsheet.batch_update(add_borders(f"{get_RandC_indexes(menesis3.cell_representation).xooo}{get_RandC_indexes(menesis3.cell_representation).oxoo}:{get_RandC_indexes(menesis3.cell_representation).ooxo}{int(get_RandC_indexes(menesis3.cell_representation).ooox) + 1 + len(sabatskolas_vaditaji)}", "+"))
    # Ievada vadītājus
    time.sleep(60) #Bišķ jāuzgaida, jo savādāk pārsniedz quota limitu
    j = 0
    while j < len(sabatskolas_vaditaji):
        sheet.spreadsheet.batch_update(merge_cells(f"A{j + 4}:B{j + 4}"))
        sheet.spreadsheet.batch_update(mainit_krasu(f"A{j + 4}:B{j + 4}", 233, 242, 250))
        sheet.format(f"A{j + 4}:B{j + 4}", {"horizontalAlignment": "CENTER"})
        sheet.spreadsheet.batch_update(add_borders(f"A{j + 4}:B{j + 4}", "+"))
        enter_values_for_range([sabatskolas_vaditaji[j].vards], [f"A{j + 4}:B{j + 4}"])
        j += 1

# Linijas klase
class Line:
    def __init__(self, cells):
        self.locationX = locationX = "T10" # Mazums gadās kļūda
        self.cells = cells
        self.datums = cells[0]
        cells.pop(0)
        self.vaditaja_vards = vaditaja_vards = ""
        for h in cells:
            #print(h)
            variable = sheet.get_values(h)
            if len(variable) != 0:
                #print("True")
                self.locationX = h
        self.vaditaja_vards = f"A{get_RandC_indexes(self.locationX).ox}:B{get_RandC_indexes(self.locationX).ox}"
    def colorize(self):
        for o in self.cells:
            sheet.spreadsheet.batch_update(mainit_krasu(o, 217, 234, 211))
        sheet.spreadsheet.batch_update(mainit_krasu(self.locationX, 182, 215, 168))
        sheet.spreadsheet.batch_update(mainit_krasu(self.datums, 182, 215, 168))
        sheet.spreadsheet.batch_update(mainit_krasu(self.vaditaja_vards, 182, 215, 168))

# Mainīgā līnija
def mainiga_linija_function(current_day, current_month):
    linijas_shunas = []
    sabati = []
    range_ = []
    index = 0
    if current_month == pirma_menesha_index:
        sabati = menesis1.all_sabaths
        range_ = get_list_of_all_cells_from_range(next_row(menesis1.cell_representation))
    elif current_month == otraa_menesha_index:
        sabati = menesis2.all_sabaths
        range_ = get_list_of_all_cells_from_range(next_row(menesis2.cell_representation))
    elif current_month == tresha_menesha_index:
        sabati = menesis3.all_sabaths
        range_ = get_list_of_all_cells_from_range(next_row(menesis3.cell_representation))
    for k in sabati: 
        if k < current_day:
            index = return_index_from_value_of_list(sabati, k + 7)
    if index != None:
        linijas_shunas = get_list_of_all_cells_from_range(f"{range_[index]}:{get_RandC_indexes(range_[index]).xo}{int(get_RandC_indexes(range_[index]).ox) + len(sabatskolas_vaditaji)}")
    else:
        collonas_burts = get_RandC_indexes(range_[len(range_) - 1]).xo
        for t in alphabet:
            if t == collonas_burts.lower():
                skaitlis = alphabet[t] + 1
        for g in alphabet:
            if alphabet[g] == skaitlis:
                collonas_burts = g
        linijas_shunas = get_list_of_all_cells_from_range(f"{collonas_burts}{get_RandC_indexes(range_[len(range_) - 1]).ox}:{collonas_burts}{int(get_RandC_indexes(range_[len(range_) - 1]).ox) + len(sabatskolas_vaditaji)}")
    linija = Line(linijas_shunas)
    return linija

def delete_previouse_mainiga_linija(linija_objekts):
    for o in linija_objekts.cells:
        sheet.spreadsheet.batch_update(mainit_krasu(o, 255, 255, 255))
    sheet.spreadsheet.batch_update(mainit_krasu(linija_objekts.locationX, 255, 255, 255))
    sheet.spreadsheet.batch_update(mainit_krasu(linija_objekts.datums, 233, 242, 250))
    sheet.spreadsheet.batch_update(mainit_krasu(linija_objekts.vaditaja_vards, 233, 242, 250))


def ievadit_datus():
    collonas = 2 #jo C no kura sākas datu ievade ir 2 alphabeta
    if pagajusa_ceturksna_dati["gads"] == year and int(aktualais_ceturksnis) == pagajusa_ceturksna_dati["ceturksnis"]:
        print("Izpildijās")
        if menesis1.number_of_sabaths + menesis2.number_of_sabaths + menesis3.number_of_sabaths == len(pagajusa_ceturksna_dati["seciba"]):
            print("Sanāca")
            for a in pagajusa_ceturksna_dati["seciba"]:
                for b in sabatskolas_vaditaji:
                    if a == b.vards:
                        #print(f"{return_an_key_from_value_of_dictionary(alphabet, collonas).capitalize()}{get_RandC_indexes(b.cell_representation).oxoo}")
                        sheet.update_acell(f"{return_an_key_from_value_of_dictionary(alphabet, collonas).capitalize()}{get_RandC_indexes(b.cell_representation).oxoo}", "x")
                        collonas += 1
        else:
            start_from = pagajusa_ceturksna_dati["seciba"][len(pagajusa_ceturksna_dati["seciba"]) - 1]
            for i in sabatskolas_vaditaji:
                if i.vards == start_from:
                    start_from = f"A{int(get_RandC_indexes(i.cell_representation).oxoo) + 1}:B{int(get_RandC_indexes(i.cell_representation).ooox) + 1}"
            r1 = int(get_RandC_indexes(start_from).oxoo)
            r2 = int(get_RandC_indexes(sabatskolas_vaditaji[len(sabatskolas_vaditaji) - 1].cell_representation).oxoo) + 1
            c1 = 2
            c2 = menesis1.number_of_sabaths + menesis2.number_of_sabaths + menesis3.number_of_sabaths + 1
            while c1 < c2:
                row_input = 0
                if r1 < r2:
                    row_input = r1
                else:
                    row_input = 4 # rindiņas sākas no jauna
                    r1 = 4
                r1 += 1
                sheet.update_acell(f"{return_an_key_from_value_of_dictionary(alphabet, c1).capitalize()}{row_input}", "x")
                c1 += 1
    else:
        #print("Else variants")
        start_from = ""
        if pagajusa_ceturksna_dati["gads"] < year and aktualais_ceturksnis == "1" or int(aktualais_ceturksnis) - 1 == pagajusa_ceturksna_dati["ceturksnis"]:
            for i in sabatskolas_vaditaji:
                if i.vards == pagajusa_ceturksna_dati["seciba"][len(pagajusa_ceturksna_dati["seciba"]) - 1]:
                    start_from = f"A{int(get_RandC_indexes(i.cell_representation).oxoo) + 1}:B{int(get_RandC_indexes(i.cell_representation).ooox) + 1}"
        else:
            start_from = "C4:D4" # Vieta kas ir pati pirmā pirmajam vadītājam un pirmajam datumam
        r1 = int(get_RandC_indexes(start_from).oxoo)
        r2 = int(get_RandC_indexes(sabatskolas_vaditaji[len(sabatskolas_vaditaji) - 1].cell_representation).oxoo) + 1
        c1 = 2
        c2 = menesis1.number_of_sabaths + menesis2.number_of_sabaths + menesis3.number_of_sabaths + 1
        while c1 < c2:
            row_input = 0
            if r1 < r2:
                row_input = r1
            else:
                row_input = 4 # rindiņas sākas no jauna
                r1 = 4
            r1 += 1
            sheet.update_acell(f"{return_an_key_from_value_of_dictionary(alphabet, c1).capitalize()}{row_input}", "x")
            c1 += 1
#clear_all()
#tabula()
#ievadit_datus()
#mainiga_linija.colorize()



def main():
    
    global day
    global month
    global year
    
    if month != 12:
        if day != 29:
            day +=1
        else:
            day = 1
            month = 1
    else:
        if day < 29:
            day += 1
        else:
            month = 1
            day = 1
            year += 1
    #(f"month is {month}")
    #print(f"Pirmā mēneša index ir {pirma_menesha_index}")
    #print(f"{day}.{month}.{year}")        

    galvenas_operacijas()
    #print(f"Pirmā mēneša index ir {pirma_menesha_index}")
    global mainiga_linija
    if year > pagajusa_ceturksna_dati["gads"] or  month > menesis3.index:
        #print(f"year({year}) > pagajusa_ceturksna_dati['gads']({pagajusa_ceturksna_dati['gads']}) OR menesis3.index({menesis3.index}) > month({month})")
        #print("Tici vai nē")
        print(pagajusa_ceturksna_dati)
        print("Updato JSON laikam")
        if pagajusa_ceturksna_dati["ceturksnis"] == 4:
            pagajusa_ceturksna_dati["ceturksnis"] = 1
            pagajusa_ceturksna_dati["gads"] = pagajusa_ceturksna_dati["gads"] + 1
            vards = ""
            #for i in mainiga_linija.cells:
                #u = sheet.get_values(i)
                #if len(u) != 0:
                    #vards = u
            vards = pagajusa_ceturksna_dati["seciba"][len(pagajusa_ceturksna_dati["seciba"]) - 1]
            #j = 0
            #while j < len(pagajusa_ceturksna_dati["seciba"]):
                #pagajusa_ceturksna_dati["seciba"].pop(j)
                #j += 1
            pagajusa_ceturksna_dati["seciba"] = []
            pagajusa_ceturksna_dati["seciba"].append(vards)
        print(pagajusa_ceturksna_dati)
        #clear_all()
        #tabula()
        ievadit_datus()
        mainiga_linija.colorize
    else:
        atzimetais_datums = sheet.get_values(mainiga_linija.datums)
        print(atzimetais_datums)
        print(atzimetais_datums[0][0])
        atzimetais_datums = int(atzimetais_datums[0][0])
        if atzimetais_datums < day:
            delete_previouse_mainiga_linija(mainiga_linija)
            mainiga_linija = mainiga_linija_function(day, month)
            mainiga_linija.colorize()
    time.sleep(60)
    #print(f"Day {day}")
    #print(f"Month {month}")
    #print(f"Year {year}")
    main()

year = 2022
month = 12
day = 30
galvenas_operacijas()
mainiga_linija = mainiga_linija_function(day, month)
main()

