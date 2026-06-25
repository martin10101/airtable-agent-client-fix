const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

const DOCX_TARGET = 'word/document.xml';

function escapeXml(s) {
  return String(s == null ? '' : s)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function decodeXml(s) {
  return String(s || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function textFromParagraph(paragraphXml) {
  const pieces = [];
  const textRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = textRegex.exec(paragraphXml)) !== null) pieces.push(decodeXml(m[1]));
  return pieces.join('');
}

function cleanText(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

function paragraphText(paragraphXml) {
  return cleanText(textFromParagraph(paragraphXml));
}

function paragraphOpen(paragraphXml) {
  return (paragraphXml.match(/^<w:p\b[^>]*>/) || ['<w:p>'])[0];
}

function paragraphProperties(paragraphXml) {
  const match = paragraphXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/);
  return match ? match[0] : '';
}

function runProperties(paragraphXml) {
  const run = paragraphXml.match(/<w:r\b[\s\S]*?<\/w:r>/);
  if (!run) return '';
  const props = run[0].match(/<w:rPr\b[\s\S]*?<\/w:rPr>/);
  return props ? props[0] : '';
}

function paragraphLike(paragraphXml, text) {
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return `${paragraphOpen(paragraphXml)}${paragraphProperties(paragraphXml)}<w:r>${runProperties(paragraphXml)}<w:t${space}>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function collectParagraphs(xml) {
  const paragraphs = [];
  const re = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    paragraphs.push({
      start: m.index,
      end: m.index + m[0].length,
      xml: m[0],
      text: paragraphText(m[0])
    });
  }
  return paragraphs;
}

function isProjectFactLine(text) {
  return (
    /^A new (multiple dwelling|mixed-use|residential|commercial)/i.test(text) ||
    /^The proposed is to construct/i.test(text) ||
    /^The proposed building contains/i.test(text) ||
    /^The project involves the proposed/i.test(text) ||
    /^The building (will aggregate|will cover|will include|will have a total)/i.test(text)
  );
}

function isValuationStart(text) {
  return /^(Transitional Assessed Valuation|Tax Class 2a,\s*2b and 2c CAP)/i.test(text);
}

function isPostCompletionStart(text) {
  return /^Post-Completion Projected Assessed Value and Tax Liability/i.test(text);
}

function summaryLineFor(text) {
  const suffixMatch = String(text || '').match(/\s+[–-]\s*(.+)$/);
  const suffix = suffixMatch ? suffixMatch[1].trim() : '';
  return suffix ? `Property Summary: {{Property_Summary}} – ${suffix}` : 'Property Summary: {{Property_Summary}}';
}

function valuationSlotForText(text, used) {
  const clean = String(text || '').trim();
  if (isValuationStart(clean) && !used.has('heading')) return ['heading', '{{Tax_Valuation_Heading}}'];
  if (/^Projected Assessed Value Increase and Phase-In/i.test(clean) && !used.has('phaseHeading')) {
    return ['phaseHeading', '{{Tax_Valuation_Phase_Heading}}'];
  }
  if (/^The New York Real Property Tax Law/i.test(clean) && !used.has('rptlParagraph')) {
    return ['rptlParagraph', '{{Tax_Valuation_RPTL_Paragraph}}'];
  }
  if (/^(We believe|We anticipate)/i.test(clean) && !used.has('phaseParagraph')) {
    return ['phaseParagraph', '{{Tax_Valuation_Phase_Paragraph}}'];
  }
  if (/^(Under the DOF|Under New York Real Property Tax Law)/i.test(clean) && !used.has('primaryParagraph')) {
    return ['primaryParagraph', '{{Tax_Valuation_Primary_Paragraph}}'];
  }
  return null;
}

function valuationPlaceholderParagraphs(range) {
  const used = new Set();
  const rendered = [];
  for (const p of range) {
    const slot = valuationSlotForText(p.text, used);
    if (!slot) continue;
    used.add(slot[0]);
    rendered.push(paragraphLike(p.xml, slot[1]));
  }
  if (!used.has('heading') && range[0]) rendered.unshift(paragraphLike(range[0].xml, '{{Tax_Valuation_Heading}}'));
  if (!used.has('primaryParagraph') && range[1]) rendered.push(paragraphLike(range[1].xml, '{{Tax_Valuation_Primary_Paragraph}}'));
  return rendered.join('');
}

function projectDetailSlotForText(text) {
  const clean = String(text || '').trim();
  if (/^A new multiple dwelling/i.test(clean)) return '{{Project_Details_New_Residential_Line}}';
  if (/^A new mixed-use/i.test(clean) || /^The proposed is to construct a new mixed-use/i.test(clean)) {
    return '{{Project_Details_New_Mixed_Use_Line}}';
  }
  if (/^The project involves the proposed/i.test(clean)) return '{{Project_Details_Renovation_Line}}';
  if (/^The building (will aggregate|will cover|will include|will have a total)/i.test(clean)) {
    return '{{Gross_Square_Feet_Line}}';
  }
  return null;
}

function projectDetailPlaceholderParagraphs(range, opts) {
  opts = opts || {};
  const rendered = [];
  const used = new Set();
  if (opts.includeAi !== false && range[0]) rendered.push(paragraphLike(range[0].xml, '{{Project_Details_AI_Line}}'));
  for (const p of range) {
    const slot = projectDetailSlotForText(p.text);
    if (!slot || used.has(slot)) continue;
    used.add(slot);
    rendered.push(paragraphLike(p.xml, slot));
  }
  return rendered.join('');
}

function shouldSmartCopy(filePath) {
  const base = path.basename(filePath);
  if (!/\.docx$/i.test(base)) return false;
  if (/^~\$/.test(base)) return false;
  if (!/^Tmplt/i.test(base)) return false;
  return /Rental Project/i.test(base) || /485x|485-x|ICAP/i.test(base);
}

function rewriteDocumentXml(xml) {
  const paragraphs = collectParagraphs(xml);
  let out = '';
  let cursor = 0;
  let changed = false;
  let insideProjectDetails = false;
  let projectBlockInserted = false;
  let summaryInserted = false;

  for (let i = 0; i < paragraphs.length;) {
    const p = paragraphs[i];
    const text = p.text;
    out += xml.slice(cursor, p.start);

    if (/^Project details$/i.test(text)) {
      insideProjectDetails = true;
      projectBlockInserted = false;
      out += p.xml;
      cursor = p.end;
      i++;
      continue;
    }

    if (/^The real estate tax benefits available/i.test(text)) {
      insideProjectDetails = false;
      out += p.xml;
      cursor = p.end;
      i++;
      continue;
    }

    if (i < 20 && /^RE\s*:/i.test(text)) {
      out += paragraphLike(p.xml, 'RE:\t{{Property_Address}}');
      cursor = p.end;
      i++;
      changed = true;
      continue;
    }

    if (i < 20 && /^[A-Za-z][A-Za-z .'-]+,\s*NY$/i.test(text)) {
      out += paragraphLike(p.xml, '{{Borough}}, NY');
      cursor = p.end;
      i++;
      changed = true;
      continue;
    }

    if (i < 25 && /^Block\s*&\s*Lot\s*:/i.test(text)) {
      out += paragraphLike(p.xml, 'Block & Lot: {{Block_Lot}}');
      cursor = p.end;
      i++;
      changed = true;
      continue;
    }

    if (/^(Project Details|Property Summary)\s*:/i.test(text)) {
      if (!summaryInserted) {
        out += paragraphLike(p.xml, summaryLineFor(text));
        summaryInserted = true;
      }
      cursor = p.end;
      i++;
      changed = true;
      continue;
    }

    if (insideProjectDetails && isProjectFactLine(text)) {
      let next = i + 1;
      while (next < paragraphs.length && isProjectFactLine(paragraphs[next].text)) next++;
      out += projectDetailPlaceholderParagraphs(paragraphs.slice(i, next), { includeAi: !projectBlockInserted });
      projectBlockInserted = true;
      cursor = paragraphs[next] ? paragraphs[next].start : p.end;
      i = next;
      changed = true;
      continue;
    }

    if (isValuationStart(text)) {
      let next = i + 1;
      while (next < paragraphs.length && !isPostCompletionStart(paragraphs[next].text)) next++;
      out += valuationPlaceholderParagraphs(paragraphs.slice(i, next));
      cursor = paragraphs[next] ? paragraphs[next].start : p.end;
      i = next;
      changed = true;
      continue;
    }

    if (/new construction\s+renovation/i.test(text)) {
      out += paragraphLike(p.xml, text.replace(/new construction\s+renovation/ig, '{{Construction_Type}}'));
      cursor = p.end;
      i++;
      changed = true;
      continue;
    }

    out += p.xml;
    cursor = p.end;
    i++;
  }

  out += xml.slice(cursor);
  return { xml: out, changed };
}

function createSmartTemplateCopy(sourcePath, outputPath) {
  const zip = new PizZip(fs.readFileSync(sourcePath));
  const doc = zip.file(DOCX_TARGET);
  if (!doc) return { ok: false, changed: false, reason: 'missing word/document.xml' };

  const result = rewriteDocumentXml(doc.asText());
  if (!result.changed) return { ok: false, changed: false, reason: 'no recognized template sections found' };

  zip.file(DOCX_TARGET, result.xml);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, zip.generate({ type: 'nodebuffer' }));
  return { ok: true, changed: true };
}

function walk(dir, results) {
  results = results || [];
  let entries = [];
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return results;
  }
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (/^(node_modules|logs|work|backup-)/i.test(entry.name)) continue;
      walk(full, results);
    } else if (entry.isFile() && shouldSmartCopy(full)) {
      results.push(full);
    }
  }
  return results;
}

function ensureSmartTemplateCopies(rootDir, opts) {
  opts = opts || {};
  const log = typeof opts.log === 'function' ? opts.log : (() => {});
  const created = [];
  const skipped = [];
  const warnings = [];
  if (!rootDir || !fs.existsSync(rootDir)) return { created, skipped, warnings: [`Template root not found: ${rootDir}`] };

  for (const sourcePath of walk(rootDir)) {
    if (sourcePath.split(/[\\/]/).some((part) => /^SMART$/i.test(part))) continue;
    const parent = path.dirname(sourcePath);
    const outputPath = path.join(parent, 'SMART', path.basename(sourcePath));
    if (fs.existsSync(outputPath)) {
      skipped.push({ sourcePath, outputPath, reason: 'smart copy already exists' });
      continue;
    }
    try {
      const result = createSmartTemplateCopy(sourcePath, outputPath);
      if (result.ok) {
        created.push({ sourcePath, outputPath });
        log(`[smart-template] Created ${outputPath}`);
      } else {
        skipped.push({ sourcePath, outputPath, reason: result.reason });
      }
    } catch (e) {
      warnings.push(`${sourcePath}: ${e.message}`);
    }
  }
  return { created, skipped, warnings };
}

module.exports = {
  createSmartTemplateCopy,
  ensureSmartTemplateCopies,
  rewriteDocumentXml,
  shouldSmartCopy
};
