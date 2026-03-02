#---------------------------------------------------------------------
# Program Name: ReadWriteGoogleSheets
#
# Purpose: This code reads from and writes to a Google Sheet
# 
# 
#
# Author: Thalia Edwards
#
# Date: 03/02/2026
#------------------------------------------------------------------------
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('pythonProject/secret_key/secret_key.json', scopes=scopes)

file = gspread.authorize(creds)
workbook = file.open('application-form-responses')
sheet = workbook.sheet1

#print every cell in the range A2:R4
for cell in sheet.range('A2:R4'):
    print(cell.value)

#print every cell in the range A:A
#print(sheet.range('A:A'))

#write to a cell
#sheet.update_acell('B2', 'Hello World')