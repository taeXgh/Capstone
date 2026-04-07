# ---------------------------------------------------------------------
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
# ------------------------------------------------------------------------
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docx import Document
from pathlib import Path
from datetime import datetime, timedelta

# --------testing-------------
import hashlib
import json
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.enum.dml import MSO_THEME_COLOR_INDEX

# ----------------------------

# username = "talop"


def get_resume_link(resume):
    return resume.split('"')[1]


def format_submission_date(raw_value):
    # Handles Google Sheets serial dates and converts to MM/DD/YYYY format
    try:
        serial = float(raw_value)
        dt = datetime(1899, 12, 30) + timedelta(days=serial)
        return dt.strftime("%m/%d/%Y")
    except (ValueError, TypeError):
        # Fallback if value is already a date string
        text = str(raw_value).strip()
        for fmt in ("%m/%d/%Y %H:%M:%S", "%m/%d/%Y"):
            try:
                return datetime.strptime(text, fmt).strftime("%m/%d/%Y")
            except ValueError:
                pass
        return text

        # Credit to ryan-rushton on GitHub for this function:


def add_hyperlink(paragraph, url, text):
    # 1) Create relationship id between this paragraph and external URL
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    # 2) Build <w:hyperlink r:id="...">
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    hyperlink.set(qn("w:history"), "1")

    # 3) Build the visible run inside the hyperlink
    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")

    # Optional style name from Word template
    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), "Hyperlink")
    r_pr.append(r_style)
    new_run.append(r_pr)

    # Add visible text
    text_elem = OxmlElement("w:t")
    text_elem.text = text
    new_run.append(text_elem)

    hyperlink.append(new_run)

    # 4) Attach hyperlink to paragraph
    run = paragraph.add_run()
    run._r.append(hyperlink)

    # 5) Visual fallback if template lacks Hyperlink style
    run.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    run.font.underline = True

    return run


scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent  # pythonProject

key_path = PROJECT_DIR / "secret_key" / "secret_key.json"
template_path = PROJECT_DIR / "Applications" / "Application Template - Copy.docx"


creds = ServiceAccountCredentials.from_json_keyfile_name(key_path, scopes=scopes)

file = gspread.authorize(creds)
workbook = file.open("application-form-responses")
sheet = workbook.sheet1


rows = sheet.get_all_values(value_render_option="FORMULA")
# -----------------------------testing-----------------------------
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
# ------------------------------------------------------------------
for row in rows[1:]:
    application = {
        "submission": format_submission_date(row[0]),
        "name": row[1],
        "email": row[2],
        "position": row[3],
        "resume_URL": get_resume_link(str(row[4])),
        "employment_status": row[5],
        "prior": row[6],
        "hear_about": row[7],
        "reason_left": row[8],
        "current_employer": row[9],
        "availability": row[10],
        "phone": row[11],
        "preferred": row[12],
        "acknowledgement": row[13],
        "age": row[14],
    }
    # ------------------------------testing-----------------------------
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

    print(
        f"Applicant ID: {applicant_id} | Position: {position} | Name: {applicant_name}"
    )
    print(application)
    print()

# Save updated registry for next script run (persists ID assignments)
with open(registry_path, "w") as f:
    json.dump(applicants_registry, f, indent=2)
# -------------------------------------------------------------------


# Process all applications and generate Word documents
for applicant_id, positions_dict in applicants.items():
    for position_name, application in positions_dict.items():
        applicant = application.get("name")
        output_path = (
            PROJECT_DIR
            / "Applications"
            / f"{applicant} - {position_name} Application.docx"
        )

        document = Document(template_path)
        for key, value in application.items():
            for p in document.paragraphs:
                placeholder = f"[{key}]"
                if placeholder in p.text:
                    if key == "resume_URL":
                        # Remove placeholder text, then insert hyperlink
                        p.text = p.text.replace(placeholder, "")
                        add_hyperlink(p, str(value), "View Resume")
                    else:
                        p.text = p.text.replace(placeholder, str(value))

        document.save(output_path)
