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

#print every cell in the range A2:R502
# for cell in sheet.range('A2:R4'):
#     if cell.value != '':
#         print(cell.value)
#     else:
#         print('No Value')
#desired output:
# Submission Date: 02/18/2026 12:00:00
# Name: John Doe
# Email: tedwar6@gmu.edu
# Subject: test
# Position: Certified Nurse Assistant
# File Upload: resume.docx
# are you currently employed: 	Yes
# do you have a nursing license: 	No
# have you ever worked with us before:	No
# how did you hear about us:	Online
# if yes why did you leave:	Empty
# if yes with who:	George Mason University
# work availability:	Part Time
# work phone:	(555) 555-5555
# authorization:	Yes
# preferred method of contact:	Email

# Get headers from row 1 (A1:R1)
headers = [h.strip() for h in sheet.row_values(1)]

# Get response rows (A2:R4)
rows = sheet.get_all_values()[1:4]   # rows 2-4

for row in rows:
    # Pad short rows so zip doesn't stop early
    padded_row = row + [''] * (len(headers) - len(row))

    for header, value in zip(headers, padded_row):
        clean_value = value.strip()
        display_value = clean_value if clean_value else "Empty"
        #print(f"{header}: {display_value}")   # f-string
        print("{:<38} {}".format(f"{header}:", display_value)) #.format for alignment
    print()  # blank line between submissions


#print every cell in the range A:A
#print(sheet.range('A:A'))

#write to a cell
#sheet.update_acell('B2', 'Hello World')