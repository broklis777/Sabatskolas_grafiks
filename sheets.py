import gspread
from oauth2client.service_account import ServiceAccountCredentials
import calendar
import time
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Sabatskolas grafiks").sheet1
#sheet.update_cell(2,2, "CHANGED")
#value = sheet.cell(2,3).value
#print(value)
