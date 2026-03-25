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
#
# Edits by: Taha Chowdhury between 3-02-26 - 3-12-2026
#------------------------------------------------------------------------
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docx import Document
from pathlib import Path


#username = "talop"

def get_resume_link(resume):
    return resume.split('"')[1]

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent   # pythonProject

key_path = PROJECT_DIR / "secret_key" / "secret_key.json"
template_path = PROJECT_DIR / "Applications" / "Application Template - Copy.docx"


creds = ServiceAccountCredentials.from_json_keyfile_name(key_path, scopes=scopes)

file = gspread.authorize(creds)
workbook = file.open('application-form-responses')
sheet = workbook.sheet1


rows = sheet.get_all_values(value_render_option = "FORMULA") 


for row in rows[1:]:
    application = {
    "submission":         row[0],
    "name":               row[1],
    "email":              row[2],
    "position":           row[3],
    "resume_URL":         get_resume_link(str(row[4])),
    "employment_status":  row[5],
    "prior" :             row[6],
    "hear_about" :        row[7],
    "reason_left" :       row[8],
    "current_employer" :  row[9],
    "availability" :      row[10],
    "phone" :             row[11],
    "preferred" :         row[12],
    "acknowledgement" :   row[13],
    "age":                row[14],
    }

    print(application)
    print()

    applicant = application.get("name")

    output_path = PROJECT_DIR / "Applications" / f"{applicant} Application.docx"
    
    document = Document(template_path)
    for key, value in application.items():
        for p in document.paragraphs:
            if key == "resume_URL":
                if f"[{key}]" in p.text:
                    p.text = p.text.replace(f"[{key}]", str(value))

            elif f"[{key}]" in p.text:
                p.text = p.text.replace(f"[{key}]", str(value))
                
    document.save(output_path)