const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const PizZip = require('pizzip');

function colLetter(colNumber) {
  let s = '';
  let n = colNumber;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function extractXlsxContent(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheetsData = {};

  workbook.eachSheet((sheet) => {
    const rows = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const cells = {};
      let hasAnyValue = false;
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const ref = `${colLetter(colNumber)}${rowNumber}`;
        let value = cell.value;
        if (value && typeof value === 'object') {
          if (value.richText) {
            value = value.richText.map((r) => r.text).join('');
          } else if (value.formula) {
            value = value.result == null ? null : value.result;
          } else if (value.text) {
            value = value.text;
          }
        }
        if ((value == null || value === '') && !cell.formula) return;
        hasAnyValue = true;
        const out = { value };
        if (cell.formula) out.hasFormula = true;
        cells[ref] = out;
      });
      if (hasAnyValue) rows.push({ row: rowNumber, cells });
    });
    sheetsData[sheet.name] = rows;
  });

  return JSON.stringify(sheetsData);
}

function normalizeCell(v) {
  if (v == null) return null;
  if (typeof v === 'object') {
    if (v.result != null) return v.result;
    if (v.text != null) return v.text;
    if (v.formula) return null;
    return null;
  }
  return v;
}

// Legacy swap mode. Kept for compatibility, but /generate uses the XML
// yellow-cell mode below for Excel so formulas are not round-tripped by ExcelJS.
async function fillXlsxSwaps(templatePath, swapsBySheet, outputPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const applied = [];
  const skippedFormula = [];
  const mismatched = [];
  const sheetNotFound = [];

  for (const [sheetName, swaps] of Object.entries(swapsBySheet || {})) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      sheetNotFound.push({ sheetName, count: (swaps || []).length });
      continue;
    }
    for (const swap of (swaps || [])) {
      if (!swap || !swap.cellRef) continue;
      const cell = sheet.getCell(swap.cellRef);
      if (cell.formula) {
        skippedFormula.push({ sheet: sheetName, ...swap });
        continue;
      }

      const current = normalizeCell(cell.value);
      const wantOld = swap.oldValue;
      const isFillingBlank = (wantOld == null || wantOld === '') && (current == null || current === '');

      if (!isFillingBlank) {
        const a = current == null ? '' : String(current).trim();
        const b = wantOld == null ? '' : String(wantOld).trim();
        const numEq =
          a !== '' && b !== '' && !isNaN(Number(a)) && !isNaN(Number(b)) &&
          Number(a) === Number(b);
        if (!(a === b || numEq)) {
          mismatched.push({ sheet: sheetName, ...swap, actualValue: current });
          continue;
        }
      }

      cell.value = swap.newValue;
      applied.push({ sheet: sheetName, ...swap });
    }
  }

  await workbook.xlsx.writeFile(outputPath);
  return { outputPath, applied, skippedFormula, mismatched, sheetNotFound };
}

async function fillXlsx(templatePath, cellUpdates, sheetName, outputPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  let sheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  if (!sheet) sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Workbook has no sheets');

  const skipped = [];
  for (const [cellRef, value] of Object.entries(cellUpdates || {})) {
    const cell = sheet.getCell(cellRef);
    if (cell.formula) { skipped.push(cellRef); continue; }
    cell.value = value;
  }

  await workbook.xlsx.writeFile(outputPath);
  return { outputPath, skippedFormulaCells: skipped, sheetUsed: sheet.name };
}

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function xmlUnescape(value) {
  return String(value || '')
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&');
}

function parseAttrs(attrText) {
  const attrs = {};
  String(attrText || '').replace(/([\w:.-]+)="([^"]*)"/g, (_m, key, value) => {
    attrs[key] = xmlUnescape(value);
    return _m;
  });
  return attrs;
}

function removeAttr(attrText, attrName) {
  const re = new RegExp(`\\s${attrName}="[^"]*"`, 'g');
  return String(attrText || '').replace(re, '');
}

function setAttr(attrText, attrName, value) {
  const without = removeAttr(attrText, attrName);
  return `${without} ${attrName}="${xmlEscape(value)}"`;
}

function normalizeKey(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function isBlank(value) {
  return value == null || String(value).trim() === '';
}

function cellRefParts(ref) {
  const m = /^([A-Z]+)(\d+)$/i.exec(String(ref || ''));
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = (col * 26) + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
}

function cellKey(row, col) {
  return `${row}:${col}`;
}

function getZipText(zip, zipPath) {
  const file = zip.file(zipPath);
  return file ? file.asText() : null;
}

function normalizeZipPath(baseDir, target) {
  const cleaned = String(target || '').replace(/\\/g, '/');
  if (cleaned.startsWith('/')) return cleaned.replace(/^\/+/, '');
  return path.posix.normalize(path.posix.join(baseDir, cleaned));
}

function parseRelationships(xml, baseDir) {
  const out = {};
  if (!xml) return out;
  xml.replace(/<Relationship\b([^>]*)\/?>/g, (_m, attrsText) => {
    const attrs = parseAttrs(attrsText);
    if (attrs.Id && attrs.Target) out[attrs.Id] = normalizeZipPath(baseDir, attrs.Target);
    return _m;
  });
  return out;
}

function parseWorkbookSheets(zip) {
  const workbookXml = getZipText(zip, 'xl/workbook.xml');
  const relsXml = getZipText(zip, 'xl/_rels/workbook.xml.rels');
  const rels = parseRelationships(relsXml, 'xl');
  const sheets = [];
  if (!workbookXml) return sheets;
  workbookXml.replace(/<sheet\b([^>]*)\/?>/g, (_m, attrsText) => {
    const attrs = parseAttrs(attrsText);
    const rid = attrs['r:id'];
    if (!attrs.name || !rid || !rels[rid]) return _m;
    sheets.push({ name: attrs.name, relId: rid, path: rels[rid] });
    return _m;
  });
  return sheets;
}

function parseSharedStrings(zip) {
  const xml = getZipText(zip, 'xl/sharedStrings.xml');
  const strings = [];
  if (!xml) return strings;
  xml.replace(/<si\b[^>]*>([\s\S]*?)<\/si>/g, (_m, inner) => {
    const parts = [];
    inner.replace(/<t\b[^>]*>([\s\S]*?)<\/t>/g, (_tm, text) => {
      parts.push(xmlUnescape(text));
      return _tm;
    });
    strings.push(parts.join(''));
    return _m;
  });
  return strings;
}

function isYellowFill(fillXml) {
  if (!fillXml) return false;
  const colorMatches = [...fillXml.matchAll(/<(?:fgColor|bgColor)\b([^>]*)\/?>/g)];
  return colorMatches.some((m) => {
    const attrs = parseAttrs(m[1]);
    const rgb = String(attrs.rgb || '').toUpperCase();
    if (rgb.endsWith('FFFF00') || rgb === 'FFFF00') return true;
    if (String(attrs.indexed || '') === '13') return true;
    return false;
  });
}

function parseYellowStyleIndexes(zip) {
  const xml = getZipText(zip, 'xl/styles.xml');
  if (!xml) return new Set();
  const fillsSection = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(xml);
  const fillIsYellow = [];
  if (fillsSection) {
    fillsSection[1].replace(/<fill\b[^>]*>([\s\S]*?)<\/fill>/g, (_m, fillXml) => {
      fillIsYellow.push(isYellowFill(fillXml));
      return _m;
    });
  }

  const yellowStyleIndexes = new Set();
  const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (cellXfs) {
    let idx = 0;
    cellXfs[1].replace(/<xf\b([^>]*?)(?:\/>|>[\s\S]*?<\/xf>)/g, (_m, attrsText) => {
      const attrs = parseAttrs(attrsText);
      const fillId = Number(attrs.fillId || 0);
      if (fillIsYellow[fillId]) yellowStyleIndexes.add(idx);
      idx += 1;
      return _m;
    });
  }
  return yellowStyleIndexes;
}

function readInlineString(inner) {
  const parts = [];
  String(inner || '').replace(/<t\b[^>]*>([\s\S]*?)<\/t>/g, (_m, text) => {
    parts.push(xmlUnescape(text));
    return _m;
  });
  return parts.join('');
}

function getCellDisplayValue(attrs, inner, sharedStrings) {
  const type = attrs.t;
  if (type === 'inlineStr') return readInlineString(inner);
  const vMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner || '');
  if (!vMatch) return '';
  const raw = xmlUnescape(vMatch[1]);
  if (type === 's') {
    const idx = Number(raw);
    return Number.isFinite(idx) && sharedStrings[idx] != null ? sharedStrings[idx] : raw;
  }
  return raw;
}

function parseSheetCells(sheetXml, sharedStrings) {
  const cells = [];
  const valuesByCoord = new Map();
  const cellRegex = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  sheetXml.replace(cellRegex, (full, attrsText, inner = '') => {
    const attrs = parseAttrs(attrsText);
    if (!attrs.r) return full;
    const parts = cellRefParts(attrs.r);
    if (!parts) return full;
    const value = getCellDisplayValue(attrs, inner, sharedStrings);
    const hasFormula = /<f\b/.test(inner);
    const styleIndex = attrs.s == null ? 0 : Number(attrs.s);
    const cell = {
      ref: attrs.r,
      row: parts.row,
      col: parts.col,
      attrsText,
      attrs,
      inner,
      value,
      hasFormula,
      styleIndex: Number.isFinite(styleIndex) ? styleIndex : 0,
      xml: full
    };
    cells.push(cell);
    if (!isBlank(value)) valuesByCoord.set(cellKey(parts.row, parts.col), value);
    return full;
  });
  return { cells, valuesByCoord };
}

function findLeftLabel(cell, valuesByCoord) {
  for (let col = cell.col - 1; col >= 1 && col >= cell.col - 8; col -= 1) {
    const value = valuesByCoord.get(cellKey(cell.row, col));
    if (!isBlank(value)) return String(value).trim();
  }
  return '';
}

function findAboveLabel(cell, valuesByCoord) {
  for (let row = cell.row - 1; row >= 1 && row >= cell.row - 6; row -= 1) {
    const value = valuesByCoord.get(cellKey(row, cell.col));
    if (!isBlank(value)) return String(value).trim();
  }
  return '';
}

function buildFieldLookup(fields) {
  const lookup = new Map();
  for (const [name, value] of Object.entries(fields || {})) {
    lookup.set(normalizeKey(name), { name, value });
  }
  return lookup;
}

function getField(fieldsLookup, names) {
  for (const name of names) {
    const hit = fieldsLookup.get(normalizeKey(name));
    if (hit && !isBlank(hit.value)) return hit;
  }
  return null;
}

function parseListField(value) {
  if (Array.isArray(value)) return value.map((v) => String(v).trim()).filter(Boolean);
  return String(value == null ? '' : value)
    .split(/[,\n;]+/)
    .map((v) => v.trim())
    .filter(Boolean);
}

function parseNumberLike(value, isPercent) {
  if (typeof value === 'number') {
    if (isPercent && Math.abs(value) > 1.5) return value / 100;
    return value;
  }
  const text = String(value == null ? '' : value).trim();
  if (!text) return null;
  const percent = /%$/.test(text);
  const cleaned = text.replace(/[$,%\s]/g, '');
  if (!cleaned || isNaN(Number(cleaned))) return null;
  let n = Number(cleaned);
  if (isPercent || percent) n = n / 100;
  return n;
}

function computedGrossSqft(fieldsLookup) {
  const direct = getField(fieldsLookup, [
    'Gross SQFT',
    'Gross SF',
    'Total Gross SQFT',
    'Total Gross SF',
    'Building Gross SQFT',
    'Building Gross SF'
  ]);
  if (direct) return direct;

  const residential = getField(fieldsLookup, ['Residential Gross SQFT', 'Residential Gross SF', 'Residential Gross Sq Ft']);
  const commercial = getField(fieldsLookup, ['Commercial Gross SQFT', 'Commercial Gross SF', 'Commercial Gross Sq Ft']);
  const r = residential ? parseNumberLike(residential.value, false) : null;
  const c = commercial ? parseNumberLike(commercial.value, false) : null;
  if (r != null || c != null) {
    return {
      name: [residential && residential.name, commercial && commercial.name].filter(Boolean).join(' + '),
      value: (r || 0) + (c || 0)
    };
  }
  return null;
}

const LABEL_RULES = [
  {
    label: 'Annual Gross Income',
    patterns: [/\bannual\s+income\b/i, /\bannual\s+gross\s+income\b/i, /\bincome\s*\(owner/i],
    fields: ['Annual Gross Income', 'Annual Income', 'Income']
  },
  {
    label: 'Units',
    patterns: [/^units?$/i, /\bdwelling\s+units?\b/i],
    fields: ['Units', 'Residential Units', 'Dwelling Units']
  },
  {
    label: 'Expense Ratio',
    patterns: [/\bexpense\s+ratio\b/i],
    fields: ['Expense Ratio (%)', 'Expense Ratio'],
    percent: true
  },
  {
    label: 'CAP',
    patterns: [/^cap\s*(?:\(%\))?$/i, /\bcap\s*\(%\)/i, /\bcap\s+rate\b/i],
    fields: ['CAP (%)', 'CAP', 'Cap Rate', 'Capitalization Rate'],
    percent: true
  },
  {
    label: 'Tax Class',
    patterns: [/\btax\s+class\b/i],
    fields: ['Tax Class']
  },
  {
    label: 'Tax Rate',
    patterns: [/\btax\s+rate\b/i],
    fields: ['Tax Rate (%)', 'Tax Rate'],
    percent: true
  },
  {
    label: 'Residential Gross SQFT',
    patterns: [
      /\btotal\s+residential\s+s\.?f\.?\b/i,
      /^residential$/i,
      /^residential\b/i,
      /\bresidential\s+(?:gross\s+)?(?:sqft|sf|square\s+feet|sq\.?\s*ft\.?)\b/i
    ],
    fields: ['Residential Gross SQFT', 'Residential Gross SF', 'Residential Gross Sq Ft', 'Total Residential SF']
  },
  {
    label: 'Commercial Gross SQFT',
    patterns: [
      /^commercial$/i,
      /^commercial\b/i,
      /\bcommercial\s+(?:gross\s+)?(?:sqft|sf|square\s+feet|sq\.?\s*ft\.?)\b/i
    ],
    fields: ['Commercial Gross SQFT', 'Commercial Gross SF', 'Commercial Gross Sq Ft']
  },
  {
    label: 'Gross SQFT',
    patterns: [/^gross\s+(?:sf|sqft|sq\.?\s*ft\.?)$/i, /\bgross\s+(?:square\s+feet|sqft|sf)\b/i],
    fields: ['Gross SQFT', 'Gross SF', 'Total Gross SQFT', 'Total Gross SF', 'Building Gross SQFT', 'Building Gross SF'],
    computed: computedGrossSqft
  },
  {
    label: 'Hard Cost',
    patterns: [/\bconstruction\s+costs?\b/i, /\bhard\s+cost\b/i],
    fields: ['Hard Cost', 'Construction Cost', 'Construction Costs']
  },
  {
    label: 'Land SF',
    patterns: [/\bland\s+sf\b/i, /\blot\s+sf\b/i],
    fields: ['Land SF', 'Lot SF', 'Land Square Feet', 'Lot Square Feet']
  },
  {
    label: 'AV PSF',
    patterns: [/\bav\s+per\s+sf\b/i, /\bav\s+psf\b/i],
    fields: ['AV PSF', 'AV per SF']
  },
  {
    label: 'Current AV',
    patterns: [/\bcurrent\s+av\b/i, /\bcurrent\s+assessed\s+value\b/i],
    fields: ['Current AV', 'Current Assessed Value', 'Assessed Value']
  },
  {
    label: 'Base AV',
    patterns: [/\bbase\s+av\b/i, /\bbase\s+assessed\s+value\b/i],
    fields: ['Base AV', 'Base Assessed Value']
  },
  {
    label: 'Taxable AV',
    patterns: [/\btaxable\s+av\b/i],
    fields: ['Taxable AV', 'Taxable Assessed Value']
  },
  {
    label: 'Average Rent per Unit',
    patterns: [/\baverage\s+rent\s+per\s+unit\b/i, /\brentome?e?tor\b/i],
    fields: ['Average Rent per Unit', 'Average Rent', 'Rentometer Average Rent', 'Rentomeetor Average Rent']
  },
  {
    label: 'Expense per SF',
    patterns: [/\bexpense\s+per\s+sf\b/i, /\bexpense\s+psf\b/i],
    fields: ['Expense per SF', 'Expense PSF', 'Guidelines Expense per SF']
  },
  {
    label: 'Land Increase',
    patterns: [/\bland\s+increase\b/i],
    fields: ['Land Increase']
  }
];

function resolveYellowCellValue(cell, fieldsLookup) {
  const candidates = [cell.leftLabel, cell.aboveLabel, cell.value]
    .filter((value) => !isBlank(value))
    .map((value) => String(value).trim());
  const labelText = candidates.join(' | ');
  for (const rule of LABEL_RULES) {
    if (!rule.patterns.some((pattern) => candidates.some((candidate) => pattern.test(candidate)))) continue;
    const hit = (rule.computed && rule.computed(fieldsLookup)) || getField(fieldsLookup, rule.fields);
    if (!hit || isBlank(hit.value)) {
      return { status: 'noValue', rule: rule.label, labelText };
    }
    return {
      status: 'fill',
      rule: rule.label,
      fieldName: hit.name,
      value: hit.value,
      labelText,
      percent: !!rule.percent
    };
  }
  return { status: 'noMatch', labelText };
}

function coerceCellValue(value, options = {}) {
  const numeric = parseNumberLike(value, !!options.percent);
  const text = String(value == null ? '' : value).trim();
  const looksNumeric = typeof value === 'number' || /^[$,\d.\s%-]+$/.test(text);
  if (numeric != null && looksNumeric) {
    return { type: 'number', value: numeric };
  }
  if (typeof value === 'number') return { type: 'number', value };
  return { type: 'string', value: String(value == null ? '' : value) };
}

function buildCellXml(cell, value, options) {
  const coerced = coerceCellValue(value, options);
  if (coerced.type === 'number' && Number.isFinite(coerced.value)) {
    const attrs = removeAttr(cell.attrsText, 't');
    return `<c${attrs}><v>${String(coerced.value)}</v></c>`;
  }
  const attrs = setAttr(cell.attrsText, 't', 'inlineStr');
  return `<c${attrs}><is><t>${xmlEscape(coerced.value)}</t></is></c>`;
}

function countFormulas(sheetXml) {
  return (String(sheetXml || '').match(/<f\b/g) || []).length;
}

function fillYellowCellsFromFields(templatePath, fields, outputPath, opts = {}) {
  const buffer = fs.readFileSync(templatePath);
  const zip = new PizZip(buffer);
  const sharedStrings = parseSharedStrings(zip);
  const yellowStyleIndexes = parseYellowStyleIndexes(zip);
  const sheets = parseWorkbookSheets(zip);
  const fieldsLookup = buildFieldLookup(fields);
  const result = {
    outputPath,
    mode: 'xlsx-yellow-fill-xml',
    templatePath,
    yellowStyleIndexes: [...yellowStyleIndexes].sort((a, b) => a - b),
    sheets: [],
    filled: [],
    skippedFormula: [],
    unmatchedYellowCells: [],
    noValue: [],
    formulaIntegrity: [],
    warnings: []
  };

  if (!sheets.length) {
    result.warnings.push('Workbook sheet list could not be read.');
  }
  if (!yellowStyleIndexes.size) {
    result.warnings.push('No yellow fill styles were found in workbook styles.');
  }

  for (const sheet of sheets) {
    const xml = getZipText(zip, sheet.path);
    if (!xml) {
      result.warnings.push(`Sheet XML not found for ${sheet.name}: ${sheet.path}`);
      continue;
    }
    const beforeFormulaCount = countFormulas(xml);
    const parsed = parseSheetCells(xml, sharedStrings);
    const yellowCells = parsed.cells.filter((cell) => yellowStyleIndexes.has(cell.styleIndex));
    const updates = new Map();
    const mergerField = getField(fieldsLookup, ['Merger/Apportionment', 'Merger Apportionment', 'Apportionment Lots', 'Merged Lots']);
    const mergerLots = mergerField ? parseListField(mergerField.value) : [];
    let mergerLotIndex = 0;

    const sheetSummary = {
      sheet: sheet.name,
      path: sheet.path,
      yellowCellsFound: yellowCells.length,
      filled: 0,
      skippedFormula: 0,
      unmatched: 0,
      noValue: 0
    };

    for (const cell of yellowCells) {
      cell.leftLabel = findLeftLabel(cell, parsed.valuesByCoord);
      cell.aboveLabel = findAboveLabel(cell, parsed.valuesByCoord);
      const baseInfo = {
        sheet: sheet.name,
        cellRef: cell.ref,
        oldValue: cell.value,
        leftLabel: cell.leftLabel,
        aboveLabel: cell.aboveLabel
      };

      if (cell.hasFormula) {
        sheetSummary.skippedFormula += 1;
        result.skippedFormula.push({ ...baseInfo, reason: 'formula cell' });
        continue;
      }

      if (/^lot$/i.test(String(cell.aboveLabel || '').trim()) && mergerLotIndex < mergerLots.length) {
        const lotValue = mergerLots[mergerLotIndex];
        mergerLotIndex += 1;
        const newXml = buildCellXml(cell, lotValue);
        updates.set(cell.ref, {
          xml: newXml,
          newValue: lotValue,
          resolved: {
            status: 'fill',
            rule: 'Merger/Apportionment Lot',
            fieldName: mergerField.name,
            value: lotValue,
            labelText: cell.aboveLabel
          }
        });
        sheetSummary.filled += 1;
        result.filled.push({
          ...baseInfo,
          fieldName: mergerField.name,
          rule: 'Merger/Apportionment Lot',
          newValue: lotValue,
          labelText: cell.aboveLabel
        });
        continue;
      }

      const resolved = resolveYellowCellValue(cell, fieldsLookup);
      if (resolved.status === 'fill') {
        const newXml = buildCellXml(cell, resolved.value, { percent: resolved.percent });
        updates.set(cell.ref, { xml: newXml, newValue: resolved.value, resolved });
        sheetSummary.filled += 1;
        result.filled.push({
          ...baseInfo,
          fieldName: resolved.fieldName,
          rule: resolved.rule,
          newValue: resolved.value,
          labelText: resolved.labelText
        });
      } else if (resolved.status === 'noValue') {
        sheetSummary.noValue += 1;
        result.noValue.push({ ...baseInfo, rule: resolved.rule, labelText: resolved.labelText, reason: 'matching Airtable field is blank or missing' });
      } else {
        sheetSummary.unmatched += 1;
        result.unmatchedYellowCells.push({ ...baseInfo, labelText: resolved.labelText, reason: 'no label rule matched' });
      }
    }

    let nextXml = xml;
    if (updates.size) {
      nextXml = xml.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (full, attrsText) => {
        const attrs = parseAttrs(attrsText);
        const update = attrs.r ? updates.get(attrs.r) : null;
        return update ? update.xml : full;
      });
      zip.file(sheet.path, nextXml);
    }

    const afterFormulaCount = countFormulas(nextXml);
    const integrity = {
      sheet: sheet.name,
      path: sheet.path,
      before: beforeFormulaCount,
      after: afterFormulaCount,
      ok: beforeFormulaCount === afterFormulaCount
    };
    result.formulaIntegrity.push(integrity);
    if (!integrity.ok) {
      result.warnings.push(`Formula count changed on ${sheet.name}: ${beforeFormulaCount} -> ${afterFormulaCount}`);
    }
    result.sheets.push(sheetSummary);
  }

  if (result.formulaIntegrity.some((entry) => !entry.ok) && opts.failOnFormulaChange !== false) {
    throw new Error('Refusing to write XLSX because formula count changed. See Excel log details.');
  }

  const outBuffer = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(outputPath, outBuffer);
  return result;
}

module.exports = {
  extractXlsxContent,
  fillXlsx,
  fillXlsxSwaps,
  fillYellowCellsFromFields
};
