const fs = require('fs');
const path = require('path');

function apiBase() {
  const baseId = process.env.AIRTABLE_BASE_ID;
  const tableName = encodeURIComponent(process.env.AIRTABLE_TABLE_NAME);
  return `https://api.airtable.com/v0/${baseId}/${tableName}`;
}

function authHeaders(extra) {
  return Object.assign(
    {
      Authorization: `Bearer ${process.env.AIRTABLE_API_KEY}`,
      'Content-Type': 'application/json'
    },
    extra || {}
  );
}

async function getRecord(recordId) {
  const url = `${apiBase()}/${recordId}`;
  const res = await fetch(url, { headers: authHeaders() });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Airtable getRecord failed (${res.status}): ${body}`);
  }
  return res.json();
}

async function downloadAttachment(attachment, destDir) {
  if (!attachment || !attachment.url) {
    throw new Error('Attachment object has no url');
  }
  const res = await fetch(attachment.url);
  if (!res.ok) {
    throw new Error(`Attachment download failed (${res.status})`);
  }
  const buffer = Buffer.from(await res.arrayBuffer());
  const filename = attachment.filename || `template_${Date.now()}`;
  fs.mkdirSync(destDir, { recursive: true });
  const fullPath = path.join(destDir, filename);
  fs.writeFileSync(fullPath, buffer);
  return { filePath: fullPath, filename };
}

// Uses Airtable's content upload endpoint (no public URL required).
// Endpoint: POST https://content.airtable.com/v0/{baseId}/{recordId}/{fieldIdOrName}/uploadAttachment
async function uploadAttachmentFromFile(recordId, fieldName, filePath) {
  const baseId = process.env.AIRTABLE_BASE_ID;
  const filename = path.basename(filePath);
  const fileBuffer = fs.readFileSync(filePath);
  const lower = filename.toLowerCase();
  let contentType;
  if (lower.endsWith('.xlsx')) {
    contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  } else if (lower.endsWith('.pdf')) {
    contentType = 'application/pdf';
  } else {
    contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  }

  const url = `https://content.airtable.com/v0/${baseId}/${recordId}/${encodeURIComponent(fieldName)}/uploadAttachment`;
  const body = {
    contentType,
    filename,
    file: fileBuffer.toString('base64')
  };

  const res = await fetch(url, {
    method: 'POST',
    headers: authHeaders(),
    body: JSON.stringify(body)
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Airtable uploadAttachment failed (${res.status}): ${text}`);
  }
  return res.json();
}

// Fallback: PATCH the record with a publicly reachable URL.
async function patchAttachmentUrl(recordId, fieldName, fileUrl, filename) {
  const url = `${apiBase()}/${recordId}`;
  const body = {
    fields: {
      [fieldName]: [{ url: fileUrl, filename: filename }]
    }
  };
  const res = await fetch(url, {
    method: 'PATCH',
    headers: authHeaders(),
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Airtable patchAttachment failed (${res.status}): ${text}`);
  }
  return res.json();
}

// ---------- Metadata API: list tables, find a field, sync single-select options ----------

async function listTables() {
  const baseId = process.env.AIRTABLE_BASE_ID;
  const url = `https://api.airtable.com/v0/meta/bases/${baseId}/tables`;
  const res = await fetch(url, { headers: authHeaders() });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Airtable listTables failed (${res.status}): ${text}`);
  }
  return res.json(); // { tables: [ { id, name, fields: [...] }, ... ] }
}

// Returns just the fields array for a given table. Used by the Word swap
// flow so Claude knows what fields Airtable tracks.
async function getTableSchema(tableName) {
  const { tables } = await listTables();
  const table = tables.find((t) => t.name === tableName || t.id === tableName);
  if (!table) {
    throw new Error(
      `Table "${tableName}" not found in base. Check AIRTABLE_TABLE_NAME. ` +
      `PAT also needs schema.bases:read scope.`
    );
  }
  return table.fields.map((f) => ({ id: f.id, name: f.name, type: f.type, options: f.options }));
}

async function findField(tableName, fieldName) {
  const { tables } = await listTables();
  const table = tables.find((t) => t.name === tableName || t.id === tableName);
  if (!table) {
    throw new Error(
      `Table "${tableName}" not found in base. Check AIRTABLE_TABLE_NAME in .env. ` +
      `PAT also needs schema.bases:read scope.`
    );
  }
  const field = table.fields.find((f) => f.name === fieldName || f.id === fieldName);
  if (!field) {
    throw new Error(
      `Field "${fieldName}" not found in table "${table.name}". ` +
      `Create it (as a Single Select) or update TEMPLATE_SELECT_FIELD in .env.`
    );
  }
  return { tableId: table.id, tableName: table.name, field };
}

// Additive-only sync: never deletes options (Airtable's API rejects removing
// options that records are using). New files become new choices; missing files
// are reported but left alone. Prune manually if you need to.
async function syncTemplateOptions(tableName, fieldName, fileList) {
  const { tableId, field } = await findField(tableName, fieldName);
  if (field.type !== 'singleSelect') {
    throw new Error(
      `Field "${fieldName}" is type "${field.type}", expected "singleSelect". ` +
      `Change the field type in Airtable.`
    );
  }
  const existing = (field.options && field.options.choices) || [];
  const existingNames = new Set(existing.map((c) => c.name));
  const fileSet = new Set(fileList);

  const toAdd = fileList.filter((name) => !existingNames.has(name));
  const orphaned = existing.filter((c) => !fileSet.has(c.name)).map((c) => c.name);

  if (!toAdd.length) {
    return { added: [], orphaned, total: existing.length, alreadySynced: true };
  }

  // Preserve all existing choices (by id) so no records lose their value.
  const newChoices = existing
    .map((c) => ({ id: c.id, name: c.name }))
    .concat(toAdd.map((name) => ({ name })));

  const baseId = process.env.AIRTABLE_BASE_ID;
  const url = `https://api.airtable.com/v0/meta/bases/${baseId}/tables/${tableId}/fields/${field.id}`;
  const res = await fetch(url, {
    method: 'PATCH',
    headers: authHeaders(),
    body: JSON.stringify({ options: { choices: newChoices } })
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(
      `Airtable syncTemplateOptions failed (${res.status}): ${text}. ` +
      `PAT needs schema.bases:write scope.`
    );
  }
  return { added: toAdd, orphaned, total: newChoices.length, alreadySynced: false };
}

// Auto-setup: ensure the table has a Template single-select field.
// Cannot create button fields (Airtable API limitation) — those have to be added
// by the user with copy-paste formulas. Verifies the output filename + attachment
// fields exist and have the right type.
async function autoSetupTable(tableNameOrId, opts) {
  opts = opts || {};
  const templateField = opts.templateSelectField || 'Template';
  const outputNameField = opts.outputNameField || 'Property Address';
  const outputField = opts.outputField || 'Draft Letter and sheet';

  const { tables } = await listTables();
  const table = tables.find((t) => t.name === tableNameOrId || t.id === tableNameOrId);
  if (!table) throw new Error(`Table "${tableNameOrId}" not found.`);

  const report = {
    tableName: table.name,
    tableId: table.id,
    created: [],
    verified: [],
    warnings: []
  };

  const existing = (name) => table.fields.find((f) => f.name === name);

  // 1. Template single-select — create if missing, verify type if present
  const tmpl = existing(templateField);
  if (!tmpl) {
    const url = `https://api.airtable.com/v0/meta/bases/${process.env.AIRTABLE_BASE_ID}/tables/${table.id}/fields`;
    const res = await fetch(url, {
      method: 'POST',
      headers: authHeaders(),
      body: JSON.stringify({
        name: templateField,
        type: 'singleSelect',
        options: { choices: [] }
      })
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Could not create "${templateField}" field: ${text}`);
    }
    report.created.push(`${templateField} (single-select)`);
  } else if (tmpl.type !== 'singleSelect') {
    report.warnings.push(`"${templateField}" exists but type is ${tmpl.type}, expected singleSelect. Fix in Airtable.`);
  } else {
    report.verified.push(`${templateField} (single-select)`);
  }

  // 2. Output filename source (read-only check)
  const nameField = existing(outputNameField);
  if (!nameField) {
    report.warnings.push(`"${outputNameField}" field missing — will fall back to record-id-based filenames.`);
  } else {
    report.verified.push(`${outputNameField} (filename source, ${nameField.type})`);
  }

  // 3. Output attachment field
  const outField = existing(outputField);
  if (!outField) {
    report.warnings.push(`"${outputField}" field missing — filled docs will only be saved locally, not attached.`);
  } else if (outField.type !== 'multipleAttachments') {
    report.warnings.push(`"${outputField}" exists but type is ${outField.type}, expected multipleAttachments.`);
  } else {
    report.verified.push(`${outputField} (attachments — filled doc lands here)`);
  }

  return report;
}

async function updateRecordField(recordId, fieldName, value) {
  const url = `${apiBase()}/${recordId}`;
  const body = { fields: { [fieldName]: value } };
  const res = await fetch(url, { method: 'PATCH', headers: authHeaders(), body: JSON.stringify(body) });
  if (!res.ok) { const t = await res.text(); throw new Error(`Airtable updateRecordField failed (${res.status}): ${t}`); }
  return res.json();
}

module.exports = {
  updateRecordField,
  getRecord,
  downloadAttachment,
  uploadAttachmentFromFile,
  patchAttachmentUrl,
  listTables,
  getTableSchema,
  findField,
  syncTemplateOptions,
  autoSetupTable
};
