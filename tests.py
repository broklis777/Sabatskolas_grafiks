import time
import calendar
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

#tests
menesis1.tests()
menesis2.tests()
menesis3.tests()