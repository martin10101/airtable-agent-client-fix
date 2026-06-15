const ExcelJS = require('exceljs');

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

// Normalize a cell value to a comparable primitive for oldValue verification.
function normalizeCell(v) {
  if (v == null) return null;
  if (typeof v === 'object') {
    if (v.result != null) return v.result;
    if (v.text != null) return v.text;
    if (v.formula) return null; // formulas we never compare
    return null;
  }
  return v;
}

// Swap mode for Excel.
// swapsBySheet = { "SheetName": [{ cellRef, fieldName, oldValue, newValue }, ...] }
// - Formula cells are always skipped.
// - If a cell's current value doesn't match the reported oldValue, the swap is
//   recorded as "mismatched" and skipped (defensive — Claude's oldValue must be real).
// - Blank-cell fills (oldValue == null/empty) are always applied if the target
//   cell is empty and not a formula.
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
        // Compare loosely (string vs number, trimmed)
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

// Back-compat — old "fill by cell reference" flow. Preserved in case any caller
// still wants single-sheet fill without the swap safety checks.
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

module.exports = { extractXlsxContent, fillXlsx, fillXlsxSwaps };
