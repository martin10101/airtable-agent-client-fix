const fs = require('fs');
const PizZip = require('pizzip');

const DOCX_TEXT_TARGETS = [
  'word/document.xml',
  'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
  'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'
];

const SMART_TAG_RE = /\{\{\s*([A-Za-z0-9 _/&().-]+?)\s*\}\}/g;

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

function tagKey(name) {
  return String(name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getRunText(runXml) {
  const pieces = [];
  const textRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = textRegex.exec(runXml)) !== null) pieces.push(decodeXml(m[1]));
  if (/<w:tab\b/.test(runXml)) pieces.push('\t');
  if (/<w:br\b/.test(runXml)) pieces.push('\n');
  return pieces.join('');
}

function getParagraphText(paragraphXml) {
  const pieces = [];
  const runRegex = /<w:r\b[\s\S]*?<\/w:r>/g;
  let m;
  while ((m = runRegex.exec(paragraphXml)) !== null) pieces.push(getRunText(m[0]));
  return pieces.join('');
}

function paragraphTextRuns(paragraphXml) {
  const runs = [];
  const textRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = textRegex.exec(paragraphXml)) !== null) {
    runs.push({ start: m.index, end: m.index + m[0].length, text: decodeXml(m[1]) });
  }
  return runs;
}

function extractParagraphProperties(paragraphXml) {
  const match = paragraphXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/);
  return match ? match[0] : '';
}

function extractFirstRunProperties(paragraphXml) {
  const run = paragraphXml.match(/<w:r\b[\s\S]*?<\/w:r>/);
  if (!run) return '';
  const props = run[0].match(/<w:rPr\b[\s\S]*?<\/w:rPr>/);
  return props ? props[0] : '';
}

function createParagraphLike(sourceParagraphXml, text) {
  const pOpen = (sourceParagraphXml.match(/^<w:p\b[^>]*>/) || ['<w:p>'])[0];
  const pPr = extractParagraphProperties(sourceParagraphXml);
  const rPr = extractFirstRunProperties(sourceParagraphXml);
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return `${pOpen}${pPr}<w:r>${rPr}<w:t${space}>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function splitBlock(value) {
  return String(value || '')
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

function cleanupInlineText(text) {
  return String(text || '')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/\s+([,.;:])/g, '$1')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')');
}

function replaceTagsInText(text, values, state) {
  return String(text || '').replace(SMART_TAG_RE, (full, rawName) => {
    const key = tagKey(rawName);
    state.seenTags.add(rawName.trim());
    if (!Object.prototype.hasOwnProperty.call(values, key)) {
      state.missingTags.add(rawName.trim());
      return full;
    }
    state.filledTags.add(rawName.trim());
    return values[key] == null ? '' : String(values[key]);
  });
}

function renderTaggedParagraph(paragraphXml, values, state) {
  const paragraphText = getParagraphText(paragraphXml);
  SMART_TAG_RE.lastIndex = 0;
  const tagMatches = [...paragraphText.matchAll(SMART_TAG_RE)];
  if (!tagMatches.length) return paragraphXml;

  const trimmed = paragraphText.trim();
  SMART_TAG_RE.lastIndex = 0;
  const fullLine = trimmed.match(/^\{\{\s*([A-Za-z0-9 _/&().-]+?)\s*\}\}$/);
  if (fullLine) {
    const rawName = fullLine[1].trim();
    const key = tagKey(rawName);
    state.seenTags.add(rawName);
    if (!Object.prototype.hasOwnProperty.call(values, key)) {
      state.missingTags.add(rawName);
      return paragraphXml;
    }
    state.filledTags.add(rawName);
    const lines = splitBlock(values[key]);
    if (!lines.length) return '';
    return lines.map((line) => createParagraphLike(paragraphXml, line)).join('');
  }

  let directXml = paragraphXml;
  let directChanged = false;
  const textRuns = paragraphTextRuns(paragraphXml);
  for (let i = textRuns.length - 1; i >= 0; i--) {
    const run = textRuns[i];
    SMART_TAG_RE.lastIndex = 0;
    if (!SMART_TAG_RE.test(run.text)) continue;
    SMART_TAG_RE.lastIndex = 0;
    const replaced = cleanupInlineText(replaceTagsInText(run.text, values, state));
    const oldXml = directXml.slice(run.start, run.end);
    const space = /^\s|\s$/.test(replaced) ? ' xml:space="preserve"' : '';
    directXml = directXml.slice(0, run.start) + `<w:t${space}>${escapeXml(replaced)}</w:t>` + directXml.slice(run.end);
    directChanged = directChanged || oldXml !== directXml.slice(run.start, run.start + oldXml.length);
  }
  if (directChanged && !SMART_TAG_RE.test(getParagraphText(directXml))) return directXml;

  const replacedText = cleanupInlineText(replaceTagsInText(paragraphText, values, state));
  return createParagraphLike(paragraphXml, replacedText);
}

function renderSmartXml(xml, values, state) {
  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => renderTaggedParagraph(paragraph, values, state));
}

function containsSmartTags(filePath) {
  const content = fs.readFileSync(filePath);
  const zip = new PizZip(content);
  for (const target of DOCX_TEXT_TARGETS) {
    const file = zip.file(target);
    if (!file) continue;
    if (SMART_TAG_RE.test(file.asText())) return true;
    SMART_TAG_RE.lastIndex = 0;
  }
  return false;
}

function inspectSmartTemplateText(text) {
  const warnings = [];
  if (/new construction\s*,?\s*renovation|new construction renovation/i.test(text)) {
    warnings.push('Template still contains combined new construction/renovation wording.');
  }
  if (/A new multiple dwelling[\s\S]{0,1200}A new mixed-use[\s\S]{0,1200}The project involves/i.test(text)) {
    warnings.push('Template still appears to contain competing Project details options.');
  }
  if ((text.match(/Projected Assessed Value Increase and Phase-In/gi) || []).length > 1) {
    warnings.push('Template still contains duplicate Projected Assessed Value Increase headings.');
  }
  if (/Tax Class 2a[\s\S]{0,1600}Transitional Assessed Valuation|Transitional Assessed Valuation[\s\S]{0,1600}Tax Class 2a/i.test(text)) {
    warnings.push('Template may still contain both transitional and 2A/2B/2C valuation sections.');
  }
  return warnings;
}

function renderSmartDocx(templatePath, values, outputPath, opts) {
  opts = opts || {};
  const content = fs.readFileSync(templatePath);
  const zip = new PizZip(content);
  const state = {
    seenTags: new Set(),
    filledTags: new Set(),
    missingTags: new Set()
  };

  for (const target of DOCX_TEXT_TARGETS) {
    const file = zip.file(target);
    if (!file) continue;
    const xml = file.asText();
    if (!SMART_TAG_RE.test(xml)) {
      SMART_TAG_RE.lastIndex = 0;
      continue;
    }
    SMART_TAG_RE.lastIndex = 0;
    zip.file(target, renderSmartXml(xml, values, state));
  }

  const outputBuffer = zip.generate({ type: 'nodebuffer' });
  fs.writeFileSync(outputPath, outputBuffer);

  return {
    outputPath,
    filledTags: [...state.filledTags].sort(),
    missingTags: [...state.missingTags].sort(),
    seenTags: [...state.seenTags].sort()
  };
}

module.exports = {
  containsSmartTags,
  inspectSmartTemplateText,
  renderSmartDocx,
  tagKey
};
