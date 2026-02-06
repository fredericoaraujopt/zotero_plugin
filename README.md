# Zotero plugin for Google Sheets

A Google Apps Script that keeps a Zotero “reading list” in sync with a Google Sheet.  
It imports your tagged Zotero items into a structured reading list, lets you track **Status** and **Notes** in Sheets, and exports your edits back to Zotero.

---

## What it does

- **Imports** all Zotero items tagged **`reading list`** into Google Sheets.
- Copies across:
  - Title (as the **Paper** column) and sets it as a **clickable hyperlink** to the best available source URL
  - Authors
  - Year
  - Theme (derived from Zotero tags, excluding status tags and “reading list”)
- Preserves **Status** and **Notes** in Sheets during imports (Zotero updates won’t overwrite your reading workflow).
- **Exports** changes from Sheets back into Zotero:
  - Updates Zotero title, URL, tags (Theme + Status + reading list), and notes.
- **Imports new Zotero notes** and appends them into your Sheets Notes (so you can keep using Sheets as your “single view” for daily reading).

---

## Installation

### 1) Download the Google Sheets template
1. Download the template sheet from this repository (or make a copy of the provided Google Sheet link if you publish one).
2. Ensure your Reading List sheet has a header row containing at least:

   - `Paper`
   - `Authors`
   - `Year`
   - `Theme`
   - `Status`
   - `Notes`

> The script detects headers case-insensitively and stores internal fields in hidden columns (Key/Hash/LinkUrl).

---

### 2) Obtain a Zotero API key
1. Go to Zotero settings on the web and create an API key.
2. Make sure the key has the permissions needed to **read and write** to your library.

> You’ll also need your **library ID** (your Zotero user ID, or group ID if you adapt the script for groups).

---

### 3) Insert your Zotero credentials into Google Apps Script
1. In Google Sheets, go to **Extensions → Apps Script**.
2. Paste the script (or open the project if you imported it).
3. Open **Project Settings → Script Properties** and add the following:

- `ZOTERO_LIBRARY_ID`  
- `ZOTERO_API_KEY`  
- `ZOTERO_INCLUDE_NOTES` (optional; `true` by default)

4. Save.

5. Run any function once (e.g. `importReadingList`) and approve the permission prompts.

---

## User manual

### 4) Tag references in Zotero
In Zotero, add the tag:

- **`reading list`**

…to any references you want included in the Sheet.

---

### 5) Import the reading list into Google Sheets
In your Google Sheet:

- Reload the spreadsheet (or reopen it)
- Use the menu: **Zotero → Import reading list from Zotero**

The script will:
- Pull all Zotero items tagged `reading list`
- Build or refresh the reading list table in Sheets
- Keep your existing **Status** and **Notes** intact

---

### 6) What gets imported (and how)
The importer fills these columns:

- **Paper**: Zotero title, saved as a hyperlink to the best available source URL  
  (prefers Zotero URL; otherwise uses DOI via `https://doi.org/...` when available)
- **Authors**
- **Year**
- **Theme**: derived from Zotero tags (excluding status tags and the `reading list` tag)

> The script also maintains a hidden **LinkUrl** column as a stable source-of-truth for hashing and hyperlink behaviour.

---

### 7) Daily workflow in Sheets
Use the Sheet as your reading dashboard:

- Update **Status** with your reading state:
  - Read / Skimmed / Priority / Not started / Not finished
- Write and maintain **Notes** while reading

Imports will not overwrite your Status or Notes.

---

### 8) Export your Sheet updates back to Zotero
When you’ve updated Notes, Status, Theme, title metadata, or the Paper hyperlink:

- Use **Zotero → Export changes to Zotero**

This will:
- Update Zotero title and URL (including hyperlink deletions)
- Update Zotero tags (Theme tags + Status tag + `reading list`)
- Export your Sheet Notes into Zotero as a “Sheets-origin” child note

If you changed core bibliographic fields (Paper/Authors/Year), you may see a confirmation warning before export.

---

### 9) Import new Zotero notes into Sheets (append-only)
If you added notes directly in Zotero and want them visible in Sheets:

- Use **Zotero → Import new Zotero notes**

This will:
- Scan the Zotero reading list for newly created notes
- Append new note snippets into your **Notes** cell in Sheets
- Mark imported notes in Zotero so they don’t get appended again

Next time you export from Sheets back to Zotero, the script will tidy organisation by removing previously-imported notes (tag/marker-based) and keeping the Sheets-origin note up to date.

---

## Notes & behaviour

- **Status + Notes are “Sheet-owned”**: imports from Zotero won’t overwrite them.
- **Paper hyperlink is “Sheet-aware”**:
  - Import sets the link based on Zotero URL/DOI
  - Export can update Zotero URL from edits you make in the Paper hyperlink (including deletion)
- **Theme options** are refreshed from Zotero tags at the start of import.

---

## Troubleshooting

- If the menu doesn’t show, reload the spreadsheet.
- If the script errors with missing config, confirm Script Properties exist:
  - `ZOTERO_LIBRARY_ID`, `ZOTERO_API_KEY`
- If hyperlinks don’t work, check the stored URL includes `https://`.
- If you use a **Zotero group library**, you’ll need to adjust API endpoints (`/groups/<id>/...`).

---

## Licence
Add whichever licence you prefer (MIT is common for scripts like this).
