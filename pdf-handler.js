const fs = require('fs');
const { PDFDocument } = require('pdf-lib');

// pdf-parse has a quirky index.js that tries to read a bundled test PDF on
// require() when module.parent is null. Importing the real file directly
// sidesteps that.
const pdfParse = require('pdf-parse/lib/pdf-parse.js');

function fieldTypeName(field) {
  const ctorName = field && field.constructor ? field.constructor.name : 'Unknown';
  return ctorName;
}

async function extractPdfContent(filePath) {
  const bytes = fs.readFileSync(filePath);
  const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });

  let fields = [];
  try {
    const form = pdfDoc.getForm();
    fields = form.getFields();
  } catch (_) {
    fields = [];
  }

  if (fields.length > 0) {
    const described = fields.map((f) => {
      const type = fieldTypeName(f);
      const info = { name: f.getName(), type };
      try {
        if (type === 'PDFTextField') {
          info.currentValue = f.getText() || '';
        } else if (type === 'PDFCheckBox') {
          info.checked = !!f.isChecked();
        } else if (type === 'PDFDropdown' || type === 'PDFOptionList') {
          info.options = f.getOptions();
          info.selected = f.getSelected();
        } else if (type === 'PDFRadioGroup') {
          info.options = f.getOptions();
          info.selected = f.getSelected();
        }
      } catch (_) {
        // Some fields throw on read — leave as-is
      }
      return info;
    });

    return {
      type: 'form',
      pageCount: pdfDoc.getPageCount(),
      fields: described
    };
  }

  // Flat PDF — extract visible text so the caller can surface it, but
  // there is nothing to fill programmatically.
  let text = '';
  let pages = pdfDoc.getPageCount();
  try {
    const parsed = await pdfParse(bytes);
    text = parsed.text || '';
    pages = parsed.numpages || pages;
  } catch (_) {
    // If pdf-parse fails (rare), still return what we have
  }

  return { type: 'text', pageCount: pages, text };
}

function coerceBool(v) {
  if (v === true || v === 1) return true;
  if (v === false || v === 0 || v == null) return false;
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === 'checked' || s === 'x' || s === '1';
}

async function fillPdfForm(templatePath, formFieldsMap, outputPath, options) {
  const opts = options || {};
  const bytes = fs.readFileSync(templatePath);
  const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const form = pdfDoc.getForm();

  const skipped = [];
  const filled = [];

  for (const [fieldName, rawValue] of Object.entries(formFieldsMap || {})) {
    let field;
    try {
      field = form.getField(fieldName);
    } catch (e) {
      skipped.push({ name: fieldName, reason: 'field not found' });
      continue;
    }

    const type = fieldTypeName(field);
    try {
      if (type === 'PDFTextField') {
        field.setText(rawValue == null ? '' : String(rawValue));
      } else if (type === 'PDFCheckBox') {
        if (coerceBool(rawValue)) field.check();
        else field.uncheck();
      } else if (type === 'PDFDropdown' || type === 'PDFOptionList') {
        field.select(String(rawValue));
      } else if (type === 'PDFRadioGroup') {
        field.select(String(rawValue));
      } else {
        skipped.push({ name: fieldName, reason: `unsupported field type: ${type}` });
        continue;
      }
      filled.push(fieldName);
    } catch (e) {
      skipped.push({ name: fieldName, reason: e.message });
    }
  }

  if (opts.flatten) {
    try {
      form.flatten();
    } catch (e) {
      // Flatten can fail on exotic forms; fall through with the unflattened copy.
    }
  }

  const outBytes = await pdfDoc.save();
  fs.writeFileSync(outputPath, outBytes);
  return { outputPath, filled, skipped };
}

module.exports = { extractPdfContent, fillPdfForm };
