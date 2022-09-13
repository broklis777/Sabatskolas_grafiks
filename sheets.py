
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import calendar
import time
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Sabatskolas grafiks").sheet1

ceturkshni = {
    "1" : ["Decembris", "Janvāris", "Februāris"],
    "2" : ["Marts", "Aprīlis", "Maijs"],
    "3" : ["Jūnijs", "Jūlijs", "Augusts"],
    "4" : ["Septembris", "Oktobris", "Novembris"],
}
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





def get_RandC_indexes (string):
    class All_Indexes:
        def __init__(self, start_column_index, end_column_index, start_row_index, end_row_index):
            self.start_column_index = start_column_index
            self.end_column_index = end_column_index
            self.start_row_index = start_row_index
            self.end_row_index = end_row_index
    rangevar = string
    start_column_index = 0
    end_column_index = 0
    start_row_index = int(rangevar[1]) - 1
    end_row_index = int(rangevar[4]) - 1
    for i in alphabet:
        if i == rangevar[0].lower():
            start_column_index = alphabet[i]
        elif i == rangevar[3].lower():
            end_column_index = alphabet[i]
    vertibas = All_Indexes(start_column_index, end_column_index, start_row_index, end_row_index)
    return vertibas




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

class Menesis:
    def __init__(self, name, all_days):
        #self.number_of_sabaths = 0
        self.all_sabaths = all_sabaths = []
        #all_sabaths = []
        self.name = name
        self.all_days = all_days
        #print(all_days) #testam
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

aktualais_ceturksnis = ""
current_menesis = ""
for u in meneshi:
    if meneshi[u] == month:
        current_menesis = u
for j in ceturkshni:
    for m in ceturkshni[j]:
        if m == current_menesis:
            aktualais_ceturksnis = j
print(current_menesis)
print(ceturkshni[aktualais_ceturksnis])

pirma_menesha_index = 0
otraa_menesha_index = 0
tresha_menesha_index = 0

for menesis_key in meneshi:
    if menesis_key == current_menesis:
        pirma_menesha_index = meneshi[menesis_key]
        otraa_menesha_index = meneshi[menesis_key] + 1
        tresha_menesha_index = meneshi[menesis_key] + 2

menesis1 = Menesis(ceturkshni[aktualais_ceturksnis][0], calendar.monthcalendar(year, pirma_menesha_index))
menesis2 = Menesis(ceturkshni[aktualais_ceturksnis][1], calendar.monthcalendar(year, otraa_menesha_index))
menesis3 = Menesis(ceturkshni[aktualais_ceturksnis][2], calendar.monthcalendar(year, tresha_menesha_index))

def return_range_in_cell_format()


mainigaa_tabula = alphabet["c"] + menesis1.number_of_sabaths + menesis2.number_of_sabaths + menesis3.number_of_sabaths
for u in alphabet:
    if alphabet[u] == mainigaa_tabula:
        mainigaa_tabula = f"C2:{u.capitalize()}9"



#tests
#menesis1.tests()
#menesis2.tests()
#menesis3.tests()

print(mainigaa_tabula)
sheet.spreadsheet.batch_update(colums_lenth_change(100, "C2:T9"))
sheet.spreadsheet.batch_update(colums_lenth_change(42, mainigaa_tabula))

