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
#--------testing-------------
import hashlib
import json
#----------------------------

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
#-----------------------------testing-----------------------------
# Load or create applicant registry (persistent ID mapping)
registry_path = PROJECT_DIR / "applicants_registry.json"
try:
    with open(registry_path, "r") as f:
        applicants_registry = json.load(f)
except FileNotFoundError:
    applicants_registry = {}  # {email: applicant_id}

# Nested dictionary to track applicants and prevent duplicates
# Structure: applicants[applicant_id][position_name] = application_data
applicants = {}
#------------------------------------------------------------------
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
#------------------------------testing-----------------------------
    # Use email as lookup key for persistent ID assignment
    email = row[2].lower().strip()
    
    # Check if we've seen this email before
    if email not in applicants_registry:
        # Assign new ID (format: APP_0001, APP_0002, etc.)
        new_id_num = len(applicants_registry) + 1
        applicant_id = f"APP_{new_id_num:04d}"
        applicants_registry[email] = applicant_id
    else:
        # Reuse existing ID for returning applicant
        applicant_id = applicants_registry[email]
    
    position = application.get("position")
    applicant_name = application.get("name")

    # Initialize applicant bucket if this is first time seeing this ID
    if applicant_id not in applicants:
        applicants[applicant_id] = {}

    # Store or overwrite application (automatically handles duplicates)
    applicants[applicant_id][position] = application

    print(f"Applicant ID: {applicant_id} | Position: {position} | Name: {applicant_name}")
    print(application)
    print()
#-------------------------------------------------------------------
#-------------------------testing-----------------------------------
# Save updated registry for next script run (persists ID assignments)
with open(registry_path, "w") as f:
    json.dump(applicants_registry, f, indent=2)
#-------------------------------------------------------------------

#Create hyperlink in Word document for resume URL
def add_hyperlink(paragraph, url, text):
    """
    A function that places a hyperlink within a paragraph object.

    :param paragraph: The paragraph we are adding the hyperlink to.
    :param url: A string containing the required url
    :param text: The text displayed for the url
    :return: A Run object containing the hyperlink
    """

    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id, )
    hyperlink.set(qn('w:history'), '1')

    # Create a w:r element
    new_run = OxmlElement('w:r')

    # Create a new w:rPr element
    rPr = OxmlElement('w:rPr')

    # Create a w:rStyle element, note this currently does not add the hyperlink style as its not in
    # the default template, I have left it here in case someone uses one that has the style in it
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'Hyperlink')

    # Join all the xml elements together add add the required text to the w:r element
    rPr.append(rStyle)
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    # Create a new Run object and add the hyperlink into it
    r = paragraph.add_run()
    r._r.append(hyperlink)

    # A workaround for the lack of a hyperlink style (doesn't go purple after using the link)
    # Delete this if using a template that has the hyperlink style in it
    r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    r.font.underline = True

    return r


# Process all applications and generate Word documents
for applicant_id, positions_dict in applicants.items():
    for position_name, application in positions_dict.items():
        applicant = application.get("name")
        output_path = PROJECT_DIR / "Applications" / f"{applicant} - {position_name} Application.docx"
        
        document = Document(template_path)
        for key, value in application.items():
            for p in document.paragraphs:
                if key == "resume_URL":
                    if f"[{key}]" in p.text:
                        p.text = p.text.replace(f"[{key}]", str(value))

                elif f"[{key}]" in p.text:
                    p.text = p.text.replace(f"[{key}]", str(value))
                    
        document.save(output_path)