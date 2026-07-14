# Automated Application Collection System

## Overview

This project is an automated pipeline designed to retrieve job application data from a Google Sheet form response workbook and generate professionally formatted Word documents (`.docx`) for each applicant based on a standard template. It manages applicant data persistently by maintaining a unique registry identifier across submissions, formats data fields (such as serial dates and embedded hyperlinks), and automatically organizes the generated files into directory structures sorted by the applied position.

The project features a simple Windows Batch interface wrapper designed for end-users (such as hiring managers or project sponsors) to run the entire extraction process with a single click.

---

## Authors & Contributors

* **Thalia Edwards** — Original Author & Lead Developer
* **Taha Chowdhury** — Code Editor & Refiner

---

## Repository Structure

```text
├── .gitignore                                 # Specifies files ignored by Git tracking
├── Collect Applications.bat                  # Batch script entry point for end-users
└── pythonProject/
    └── readGoogleSheets/
        ├── ReadWriteGoogleSheets.py           # Core application parsing and automation script
        ├── applicants_registry.json           # Persistent applicant-to-ID mappings (Git ignored)
        └── Applications/                      # Templates and workspace
            └── Application Template.docx      # Word document template with data placeholders

```

---

## Core Features

### 1. Simple User Interface

* **`Collect Applications.bat`**: Provides a straightforward terminal execution flow greeting the client sponsor ("Mrs. Simpson"), triggering the background data compilation script, and outputting execution status updates directly to the console.

### 2. Google Sheets API Integration

* Authenticates securely using OAuth2 Service Account Credentials via `gspread` and `oauth2client`.
* Extracts complete row records including raw formulas, ensuring embedded hyperlink references (such as resume links) can be retrieved and parsed programmatically instead of losing data to display text.

### 3. Persistent Registry Management

* Preserves data consistency across multiple execution cycles using `applicants_registry.json`.
* Assigns unique, sequential identifiers (`APP_0001`, `APP_0002`, etc.) mapped to a normalized representation of each applicant's email address.
* Intelligently avoids duplication by updating records if an applicant applies for the same position again.

### 4. Advanced Word Document Automation

* Parses data fields natively, converting Google Sheets numerical serial date values into standard `MM/DD/YYYY` formats.
* Dynamically scans table cell elements inside `Application Template.docx` for matching text field placeholders (e.g., `[name]`, `[email]`, `[position]`) and replaces them with corresponding sheet values.
* Constructively inserts custom, clickable Word hyperlink nodes using `OxmlElement` manipulation to handle resume attachments directly inside cells.
*Appends the applicant's name to the page section footer dynamically.

### 5. Automated File & Error Handling

* Dynamically generates regional output folders sorted by position names inside the target output path (`D:\Applications\<Position Name>\`) if they do not yet exist.
* Built-in file exception protection intercepts standard runtime blocks, warning users to close active document previews (e.g., Microsoft Word open locks) without crashing the application.

---

## Prerequisites & Installation

### Dependencies

Ensure you have Python installed along with the required libraries. Install them via `pip`:

```bash
pip install gspread oauth2client python-docx

```

### Google API Credentials Setup

1. Enable the **Google Sheets API** and **Google Drive API** in your Google Cloud Console.
2. Create a Service Account, generate a JSON credential key, and rename it to `secret_key.json`.
3. Create a directory named `secret_key` inside `pythonProject/` and place the key file inside it.
4. Share your target Google Sheet (`application-form-responses`) with the client email address listed inside your service account JSON file.

---

## How to Run

1. Open your workspace environment or navigate to the project directory.
2. Ensure your template file is present at `pythonProject/Applications/Application Template.docx`.
3. Double-click or execute **`Collect Applications.bat`** from the command line.
4. Follow the on-screen prompt by pressing **Enter** to kick off the synchronization process.
5. Processed documents will be cleanly formatted and populated under your configured system drive path at `D:\Applications\`.
