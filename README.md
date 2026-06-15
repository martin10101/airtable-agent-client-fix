# Airtable Document Agent

Local AI agent that fills Word / Excel / PDF templates with data from an
Airtable record. Designed for 100+ templates kept on your PC, picked inside
Airtable via a single-select field.

**How it feels from Airtable:**

1. You're on a record. You've already set its **Template** (a single-select
   field that lists every file in your local templates folder).
2. Click **"Run Agent"** — a browser tab flashes open, the agent reads the
   record's Template, fills the document, and uploads it back to the record's
   **Completed Document** attachment field.
3. 15–45 seconds later the tab shows Download + Show in Explorer. The
   filled file is also saved locally in `output\`, named after the record's
   **Address**.

Two supporting buttons:

- **"Open Templates"** — pops the templates folder in Windows Explorer so
  you can add, rename, or organize files (subfolders are supported).
- **"Refresh Templates"** — after you add or remove files, click this to
  sync the folder contents into the Airtable **Template** single-select
  field, so every record's Template dropdown sees the new files.

---

## 1. First-time setup (client PC)

Everything is double-click. No Command Prompt needed unless something fails.

### 1a. Copy the folder

Copy the whole `airtable-agent` folder to somewhere you have write access,
e.g. `C:\Users\<you>\airtable-agent`. **Not** Program Files.

The folder must contain:

```
airtable-agent\
├─ server.js
├─ agent.js
├─ airtable.js
├─ docx-handler.js
├─ xlsx-handler.js
├─ pdf-handler.js
├─ index.html
├─ airtable-buttons.md         <- how to create the two Airtable buttons
├─ package.json
├─ .npmrc
├─ .env.example
├─ check-environment.bat
├─ check-connectivity.js
├─ first-time-setup.bat
├─ start-server.bat
├─ templates\                  <- drop your .docx / .xlsx / .pdf templates here
└─ README.md
```

### 1b. Install Node.js (once per PC)

If Node.js isn't already installed, grab **Node 22 LTS** from
https://nodejs.org. Defaults are fine. Close any open Command Prompt
windows afterward so the new PATH takes effect.

### 1c. Run the diagnostic

Double-click **`check-environment.bat`**. ~5 seconds. It checks Node, npm,
write permission, and whether the PC can reach the 4 services the agent
needs (`registry.npmjs.org`, `api.anthropic.com`, `api.airtable.com`,
`content.airtable.com`). If anything shows `[FAIL]`, the message tells you
exactly what to ask IT to whitelist.

### 1d. Run setup

Double-click **`first-time-setup.bat`**. Runs `npm install` with a
project-local `.npmrc` that tolerates corporate SSL inspection; creates
`.env` from the template if one doesn't exist yet. 1–3 minutes.

### 1e. Edit `.env`

Open `.env` in Notepad and fill in:

```
ANTHROPIC_API_KEY=REPLACE_WITH_ANTHROPIC_KEY
AIRTABLE_API_KEY=REPLACE_WITH_AIRTABLE_PAT
AIRTABLE_BASE_ID=appXXXXXXXXXXXXXX
AIRTABLE_TABLE_NAME=Your Table Name
OUTPUT_NAME_FIELD=Address
TEMPLATE_SELECT_FIELD=Template
TEMPLATES_FOLDER=
```

- `OUTPUT_NAME_FIELD` — which record field becomes the output filename.
  Default is `Address`. Change to `Project Name`, `Property`, etc. as fits
  your base. If the field is missing on a record, the agent falls back to
  `<template>_<recordId>`.
- `TEMPLATE_SELECT_FIELD` — name of the Airtable **Single Select** field
  that lists your templates. The "Refresh Templates" button populates it
  from your local folder. Default: `Template`.
- `TEMPLATES_FOLDER` — leave blank to use the `templates\` subfolder next to
  `server.js`. Or point to any folder on this PC (e.g.
  `C:\Users\<you>\OneDrive\Work\Templates` or `H:\shared\templates`).
  Subfolders work — a file at `<templates>/421a/cover.docx` becomes the
  option `421a/cover.docx` in Airtable.

**Airtable PAT (API key) scopes needed:**

- `data.records:read` — read records
- `data.records:write` — upload filled docs as attachments
- `schema.bases:read` — list tables/fields (for the sync)
- `schema.bases:write` — add options to the Template single-select field

Scope the PAT to your specific base. If `/sync-templates` returns a 403,
you're missing `schema.bases:write`.

### 1f. Drop your templates into the templates folder

Put any `.docx`, `.xlsx`, or `.pdf` templates into the folder referenced by
`TEMPLATES_FOLDER` (or `templates\` next to `server.js` if you left that
blank). Placeholder conventions:

- **Word (.docx):** `[Project Address]`, `[Block & Lot Number]`, etc.
- **Excel (.xlsx):** leave input cells blank; the agent fills them. Formulas
  are never overwritten.
- **PDF (.pdf):** must be a fillable AcroForm (see "PDF support" below).

Recommended Word markers for new/clean templates:

- Use `[[FIELD NAME]]` placeholders for direct Airtable values, for example
  `[[Property Address]]`, `[[Units]]`, `[[Residential Gross SQFT]]`, or
  `[[PROJECT_SUMMARY]]`.
- Wrap alternate sections with deterministic conditions:

```
[[KEEP_IF_PERMIT=NB]]
A new mixed-use building will be constructed...
[[END]]

[[KEEP_IF_PERMIT=ALT]]
The project involves alteration and conversion...
[[END]]

[[KEEP_IF_UNITS>10]]
Large-building transitional language...
[[END]]

[[KEEP_IF_UNITS<=10]]
Capped 2A/2B/2C language...
[[END]]

[[KEEP_IF_HAS_COMMERCIAL=Yes]]
Commercial language...
[[END]]
```

The app removes the marker lines and keeps only the section matching Airtable.
Old unmarked templates still work through the AI swap flow, but marked templates
are more predictable.

### 1g. Create the Airtable field + buttons

See **`airtable-buttons.md`** for exact formulas. TL;DR — add to your table:

- **Template** — *Single select* field (leave the options empty; the agent
  fills it on first Refresh).
- **Address** — single-line text (used for output filename).
- **Completed Document** — Attachment field (the agent uploads here).

Then three button fields:

- **Open Templates** → `http://localhost:3000/open-templates`
- **Refresh Templates** → `http://localhost:3000/sync-templates`
- **Run Agent** → `CONCATENATE("http://localhost:3000/?recordId=", RECORD_ID(), "&auto=1")`

### 1h. First sync

Start the server (`start-server.bat`), then click **Refresh Templates** in
Airtable once. Every file in your templates folder becomes an option in
the Template single-select. Now pick a Template on a record and click
**Run Agent**.

---

## 2. Running the agent

### Start the server

Double-click **`start-server.bat`**. A black window opens showing:

```
  Airtable Document Agent
  ------------------------
  UI:         http://localhost:3000
  Templates:  C:\Users\<you>\airtable-agent\templates
  Output:     C:\Users\<you>\airtable-agent\output
  Name by:    Airtable field "Address"
```

Keep that window open. Closing it stops the server.

### Daily flow

1. In Airtable, open a record. Set its **Template** field (single select).
2. Click **"Run Agent"**. A tab opens, the agent runs automatically.
3. 15–45 s later the tab shows **Download** and **Show in Explorer**. The
   filled doc is also uploaded back to the record's Completed Document
   attachment field.

When you add or rename template files on disk:

1. Click **"Open Templates"** — Windows Explorer opens on the folder. Drop
   in, rename, or organize (subfolders are fine).
2. Back in Airtable, click **"Refresh Templates"** once. The Template
   single-select updates to match the folder.

---

## 3. PDF support — what works and what doesn't

**Works:** PDFs with fillable AcroForm fields. This covers most government
submission forms (NYC 421a, ICAP, etc.). Text fields, checkboxes, dropdowns,
radio groups, and option lists are all supported. Optional `?flatten=true`
bakes the values in (non-editable) — useful right before submission.

**Doesn't work:** flat (scanned or non-fillable) PDFs. The agent detects
these and refuses with a clear error rather than producing unreliable
output. Two fixes:

- **Easiest (Adobe Acrobat):** open the PDF in Acrobat → **Tools** →
  **Prepare Form**. Acrobat auto-detects where fields should go. Rename
  each field to something descriptive (e.g. `Project_Address`, `Block_Lot`)
  so Claude can match them to record data. Save. Drop the fillable copy
  into your templates folder.
- Or convert the template to `.docx` with bracket placeholders.

**Tip:** meaningful field names like `Project_Address`, `Block_And_Lot`,
`Total_GSF` match semantically. Avoid generic names like `Text1`, `Text2`.

---

## 4. How it works (short version)

1. Airtable button opens `http://localhost:3000/?recordId=recXXX&auto=1`.
2. UI calls `/record-info/recXXX` → reads the record's **Template** value.
3. UI fetches `/templates`, pre-selects that template, auto-fires Generate.
4. `POST /generate { recordId, templateFilename }` runs.
5. Server fetches the record from Airtable.
6. Reads the local template:
   - `.docx` → extracts text → Claude returns `{ "[Placeholder]": "value" }` → server merges split XML runs and rewrites `word/document.xml`
   - `.xlsx` → sends workbook structure → Claude returns `{ "B3": 6005, ... }` → server writes, never overwriting formulas
   - `.pdf` → sends AcroForm field list → Claude returns `{ "Field_Name": "value" }` → server fills with `pdf-lib`
7. Output is named after `OUTPUT_NAME_FIELD` (default `Address`), saved to
   `output\`, and best-effort uploaded back to the record's
   `Completed Document` attachment field.

**Refresh Templates flow** (`GET /sync-templates`):

1. Server recursively walks `TEMPLATES_FOLDER` for `.docx`/`.xlsx`/`.pdf`.
2. Calls Airtable's metadata API: `PATCH /meta/bases/{id}/tables/{id}/fields/{id}`
   with the merged option list — existing options preserved by id, new files
   added as new choices.
3. Never deletes options (Airtable's API rejects removing an option any
   record is using). Orphans are reported on the confirmation page for you
   to clean up manually.

No hardcoded field mapping — Claude figures out which record field matches
which template placeholder by semantic meaning.

---

## 5. Auto-start on login (optional)

To run the agent automatically every time the PC boots:

1. Press `Win + R`, type `shell:startup`, Enter.
2. Right-click in the folder that opens → **New** → **Shortcut** →
   browse to `start-server.bat` → finish.

The server will be running silently in the background whenever the user is
logged in.

---

## 6. Troubleshooting

**"Cannot find module 'express'"**
You skipped first-time-setup.bat. Double-click it.

**`check-environment.bat` shows an endpoint as [FAIL]**
Send the output to your IT team. Ask them to allow outbound HTTPS to the
failed host(s). The agent needs `api.anthropic.com`, `api.airtable.com`,
and `content.airtable.com` at minimum; `registry.npmjs.org` is needed once,
for the initial `npm install`.

**`first-time-setup.bat` says "unable to verify the first certificate"**
Corporate SSL inspection. The project's `.npmrc` already sets
`strict-ssl=false`, which usually solves it. If not, ask IT to whitelist
`registry.npmjs.org`.

**Template dropdown says "(no templates found)"**
Click **Open folder**, drop a `.docx`/`.xlsx`/`.pdf` into it, then click the
↻ refresh button next to the dropdown.

**Generated document placeholders not replaced**
Word sometimes splits a placeholder like `[Project Address]` across runs
that have different formatting. The agent tries to merge them, but if
individual words in the placeholder were formatted differently it may
miss. Fix: select the placeholder in Word → Ctrl+Space to clear formatting
→ save.

**Excel cells with formulas aren't being filled**
That's intentional — the agent never overwrites formulas. Look in
`logs\server.log` for `skippedFormulaCells`.

**"This PDF has no fillable form fields"**
Open the PDF in Adobe Acrobat → **Prepare Form** → name the fields → save →
re-drop into your templates folder.

**"Could not reach agent" when clicking the Airtable button**
The `start-server.bat` window isn't running. Launch it.

**Output file has a weird name like `template_recXXX__...`**
The record didn't have a value in the `OUTPUT_NAME_FIELD` field (default
`Address`). Either fill in that field, or change `OUTPUT_NAME_FIELD` in
`.env` to a field that's always populated.

**Logs**
All activity is in `logs\server.log`.

---

## 7. File locations

- `templates\` — your blank templates. Edit freely.
- `output\` — every completed document, named by address. Safe to delete old ones.
- `work\` — temp downloads. Safe to delete.
- `logs\server.log` — request log.
- `.env` — **your API keys. Keep this file private. Do NOT share it.**
