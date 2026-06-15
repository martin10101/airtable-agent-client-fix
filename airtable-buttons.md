# Airtable setup — one field + three buttons

For 100+ templates you pick the template **inside Airtable** via a
single-select field. The local agent syncs the folder into that field, and
the "Run Agent" button reads it back to fill the document.

## One-time field setup

Add these fields to your table:

1. **Template** — *Single select* (empty to start; the agent populates it).
2. **Address** — single-line text. Used for the output filename.
3. **Completed Document** — Attachment. The agent uploads the filled doc here.

That's it. Do not pre-create options for Template — the agent will add one
per file in your local folder.

---

## Button 1 — "Open Templates"

Opens the Windows folder where your blank templates live so you can
browse, add, rename, or organize them (subfolders are supported).

- Button field type: **Open URL**
- URL (static):

```
http://localhost:3000/open-templates
```

When clicked, the tab opens, Windows Explorer pops to the folder, and the
tab self-closes after ~2 seconds.

---

## Button 2 — "Refresh Templates"

After you add/rename/remove a file, click this to push the file list into
the **Template** single-select field. Then the Template dropdown in every
record shows the new file.

- Button field type: **Open URL**
- URL (static):

```
http://localhost:3000/sync-templates
```

Shows a page listing what was added and what's orphaned (options in
Airtable with no matching file). Airtable's API won't let us delete options
that records are using, so orphan cleanup is manual.

---

## Button 3 — "Run Agent"

On a record where you've set **Template**, click this to fill the template
with that record's data. Works best with `auto=1` so it fires without
needing to click Generate in the UI.

- Button field type: **Open URL**
- URL (formula):

```
CONCATENATE(
  "http://localhost:3000/?recordId=", RECORD_ID(),
  "&auto=1"
)
```

What happens:

1. Browser opens the agent UI with the record ID filled.
2. UI calls `/record-info/<recordId>`, reads your **Template** value,
   auto-selects it in the dropdown.
3. Generate fires automatically. 15–45 s later you see Download and
   Show in Explorer buttons.
4. The filled doc is also uploaded back to **Completed Document**.

If you'd rather *not* auto-fire (i.e. open the UI so the user can change
the template first), drop `&auto=1`:

```
CONCATENATE("http://localhost:3000/?recordId=", RECORD_ID())
```

---

## Optional — force a specific template regardless of the record

If one table always uses the same template, override the record's Template:

```
CONCATENATE(
  "http://localhost:3000/?recordId=", RECORD_ID(),
  "&template=", ENCODE_URL_COMPONENT("421a/cover_letter.docx"),
  "&auto=1"
)
```

The URL's `template=` wins over the record's **Template** field.

---

## PAT (API key) scopes required

The agent calls Airtable's metadata API to sync options. Your Personal
Access Token must include:

- `data.records:read` — read records
- `data.records:write` — upload the filled doc as an attachment
- `schema.bases:read` — list tables/fields
- `schema.bases:write` — add options to the Template single-select field

Scope it to the specific base. If `/sync-templates` returns a 403, you're
missing `schema.bases:write`.

---

## Notes

- Buttons only work on a PC where `start-server.bat` is running.
- Buttons are per-PC — each analyst points at their own `localhost:3000`.
- Subfolders work: a file at `<templates>/421a/cover.docx` becomes
  option `421a/cover.docx` in the Template dropdown.
- If a record's Template refers to a file that's no longer in the folder,
  the agent fails with a clear message instead of guessing.
