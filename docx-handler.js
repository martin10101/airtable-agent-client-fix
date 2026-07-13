const PizZip = require('pizzip');
const fs = require('fs');

const DOCX_TEXT_TARGETS = [
  'word/document.xml',
  'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
  'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'
];

function escapeXml(s) {
  return sanitizeXmlText(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function sanitizeXmlText(s) {
  return String(s)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/g, '')
    .replace(/(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '$1');
}

function xmlExcerpt(xml, index) {
  const start = Math.max(0, index - 50);
  const end = Math.min(xml.length, index + 50);
  return xml
    .slice(start, end)
    .replace(/\s+/g, ' ')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '?');
}

function findIllegalXmlCharacters(xml, partName) {
  const problems = [];
  for (let i = 0; i < xml.length; i++) {
    const code = xml.charCodeAt(i);
    let type = '';
    if ((code >= 0x00 && code <= 0x08) || code === 0x0B || code === 0x0C || (code >= 0x0E && code <= 0x1F)) {
      type = 'illegal XML control character';
    } else if (code >= 0xD800 && code <= 0xDBFF) {
      const next = xml.charCodeAt(i + 1);
      if (next >= 0xDC00 && next <= 0xDFFF) {
        i++;
        continue;
      }
      type = 'unpaired high surrogate';
    } else if (code >= 0xDC00 && code <= 0xDFFF) {
      type = 'unpaired low surrogate';
    }
    if (!type) continue;
    problems.push({
      part: partName,
      type,
      index: i,
      code: `U+${code.toString(16).toUpperCase().padStart(4, '0')}`,
      excerpt: xmlExcerpt(xml, i)
    });
    if (problems.length >= 10) break;
  }
  return problems;
}

function findInvalidXmlEntities(xml, partName) {
  const problems = [];
  const entityRe = /&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);)/g;
  let match;
  while ((match = entityRe.exec(xml)) !== null) {
    problems.push({
      part: partName,
      type: 'invalid XML entity',
      index: match.index,
      excerpt: xmlExcerpt(xml, match.index)
    });
    if (problems.length >= 10) break;
  }
  return problems;
}

function getDomParser() {
  try {
    return require('@xmldom/xmldom').DOMParser;
  } catch (_) {
    return null;
  }
}

function validateXmlWellFormed(xml, partName) {
  const DOMParser = getDomParser();
  if (!DOMParser) return [];
  const problems = [];
  const parser = new DOMParser({
    onError: (level, message) => {
      problems.push({
        part: partName,
        type: level === 'fatalError'
          ? 'xml fatal parse error'
          : (level === 'warning' ? 'xml parse warning' : 'xml parse error'),
        message: String(message)
      });
    }
  });
  try {
    parser.parseFromString(xml, 'application/xml');
  } catch (e) {
    problems.push({ part: partName, type: 'xml parse exception', message: e.message });
  }
  return problems.slice(0, 10);
}

function detectParagraphPropertyOrderIssues(xml, partName) {
  const problems = [];
  const pPrRe = /<w:pPr\b[\s\S]*?<\/w:pPr>/g;
  let match;
  while ((match = pPrRe.exec(xml)) !== null) {
    const pPr = match[0];
    const keep = Math.min(
      ...['<w:keepNext', '<w:keepLines'].map((tag) => {
        const idx = pPr.indexOf(tag);
        return idx === -1 ? Number.POSITIVE_INFINITY : idx;
      })
    );
    if (!Number.isFinite(keep)) continue;
    const lateTags = ['<w:contextualSpacing', '<w:spacing', '<w:ind', '<w:jc', '<w:rPr', '<w:sectPr'];
    const lateBeforeKeep = lateTags.find((tag) => {
      const idx = pPr.indexOf(tag);
      return idx !== -1 && idx < keep;
    });
    if (!lateBeforeKeep) continue;
    problems.push({
      part: partName,
      type: 'word paragraph-property order',
      message: `${lateBeforeKeep} appears before keepNext/keepLines`,
      index: match.index,
      excerpt: xmlExcerpt(xml, match.index)
    });
    if (problems.length >= 10) break;
  }
  return problems;
}

function normalizeParagraphPropertyOrder(xml, log) {
  let changed = 0;
  const out = xml.replace(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/g, (pPr) => {
    let keepNext = '';
    let keepLines = '';
    let touched = false;
    let cleaned = pPr
      .replace(/<w:keepNext\b[^>]*(?:\/>|>[\s\S]*?<\/w:keepNext>)/g, (tag) => {
        if (!keepNext) keepNext = tag;
        touched = true;
        return '';
      })
      .replace(/<w:keepLines\b[^>]*(?:\/>|>[\s\S]*?<\/w:keepLines>)/g, (tag) => {
        if (!keepLines) keepLines = tag;
        touched = true;
        return '';
      });

    if (!touched) return pPr;
    const keepTags = `${keepNext}${keepLines}`;
    let normalized;
    const pStyleRe = /(<w:pPr\b[^>]*>\s*<w:pStyle\b[^>]*(?:\/>|>[\s\S]*?<\/w:pStyle>)\s*)/;
    if (pStyleRe.test(cleaned)) {
      normalized = cleaned.replace(pStyleRe, `$1${keepTags}`);
    } else {
      normalized = cleaned.replace(/(<w:pPr\b[^>]*>)/, `$1${keepTags}`);
    }

    if (normalized !== pPr) changed++;
    return normalized;
  });

  if (changed) {
    const dlog = typeof log === 'function' ? log : (() => {});
    dlog(`[docx-fix] Normalized paragraph keepNext/keepLines order in ${changed} paragraph(s)`);
  }
  return out;
}

function validateDocx(filePath) {
  const result = { ok: true, checkedParts: 0, problems: [] };
  let zip;
  try {
    zip = new PizZip(fs.readFileSync(filePath));
  } catch (e) {
    return {
      ok: false,
      checkedParts: 0,
      problems: [{ part: '(zip)', type: 'docx zip read error', message: e.message }]
    };
  }

  for (const name of Object.keys(zip.files)) {
    const entry = zip.files[name];
    if (!entry || entry.dir || !/\.(xml|rels)$/i.test(name)) continue;
    let xml = '';
    try {
      xml = entry.asText();
    } catch (e) {
      result.problems.push({ part: name, type: 'xml read error', message: e.message });
      continue;
    }
    result.checkedParts++;
    result.problems.push(...findIllegalXmlCharacters(xml, name));
    result.problems.push(...findInvalidXmlEntities(xml, name));
    result.problems.push(...validateXmlWellFormed(xml, name));
    if (/^word\/.*\.xml$/i.test(name)) {
      result.problems.push(...detectParagraphPropertyOrderIssues(xml, name));
    }
    if (result.problems.length >= 50) {
      result.problems = result.problems.slice(0, 50);
      break;
    }
  }
  result.ok = result.problems.length === 0;
  return result;
}

function decodeXml(s) {
  return String(s || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function escapeRegex(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function removeParagraphContaining(xml, oldValueEsc) {
  const idx = xml.indexOf(oldValueEsc);
  if (idx === -1) return null;
  let pStart = -1;
  let searchFrom = idx;
  while (searchFrom > 0) {
    const candidate = xml.lastIndexOf('<w:p', searchFrom);
    if (candidate === -1) return null;
    const next = xml[candidate + 4];
    if (next === ' ' || next === '>' || next === '/') { pStart = candidate; break; }
    searchFrom = candidate - 1;
  }
  if (pStart === -1) return null;
  const pEnd = xml.indexOf('</w:p>', idx);
  if (pEnd === -1) return null;
  return xml.slice(0, pStart) + xml.slice(pEnd + '</w:p>'.length);
}

function extractDocxText(filePath) {
  const content = fs.readFileSync(filePath);
  const zip = new PizZip(content);
  const xml = zip.file('word/document.xml').asText();

  const plain = xml
    .replace(/<w:p[^>]*>/g, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/\s+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  // Un-escape XML entities so Claude sees the literal characters.
  return decodeXml(plain);
}

function extractRunText(runXml) {
  const pieces = [];
  const textRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = textRegex.exec(runXml)) !== null) pieces.push(decodeXml(m[1]));
  if (/<w:tab\b/.test(runXml)) pieces.push('\t');
  if (/<w:br\b/.test(runXml)) pieces.push('\n');
  return pieces.join('');
}

function getRunMarker(runXml) {
  const propsMatch = runXml.match(/<w:rPr\b[\s\S]*?<\/w:rPr>/);
  const props = propsMatch ? propsMatch[0] : '';
  const highlightMatch = props.match(/<w:highlight\b[^>]*w:val="([^"]+)"/i);
  const colorMatch = props.match(/<w:color\b[^>]*w:val="([^"]+)"/i);
  const highlight = highlightMatch ? highlightMatch[1] : '';
  const color = colorMatch ? colorMatch[1] : '';
  const isYellowHighlight = /^yellow$/i.test(highlight);
  const isRedText = /^(ff0000|red|c00000|e00000)$/i.test(color);
  if (!isYellowHighlight && !isRedText) return null;
  return {
    yellowHighlight: isYellowHighlight,
    redText: isRedText,
    highlight,
    color
  };
}

function extractMarkedTargetsFromXml(xml, source) {
  const targets = [];
  xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    let current = null;
    const flush = () => {
      if (!current) return;
      const text = current.text.replace(/\s+/g, ' ').trim();
      if (text) targets.push({ source, text, yellowHighlight: current.yellowHighlight, redText: current.redText });
      current = null;
    };

    const runRegex = /<w:r\b[\s\S]*?<\/w:r>/g;
    let m;
    while ((m = runRegex.exec(paragraph)) !== null) {
      const run = m[0];
      const text = extractRunText(run);
      const marker = getRunMarker(run);
      if (!text) continue;
      if (!marker) {
        flush();
        continue;
      }
      if (!current) current = { text: '', yellowHighlight: false, redText: false };
      current.text += text;
      current.yellowHighlight = current.yellowHighlight || marker.yellowHighlight;
      current.redText = current.redText || marker.redText;
    }
    flush();
    return paragraph;
  });
  return targets;
}

function extractParagraphText(paragraphXml) {
  const pieces = [];
  const runRegex = /<w:r\b[\s\S]*?<\/w:r>/g;
  let m;
  while ((m = runRegex.exec(paragraphXml)) !== null) {
    pieces.push(extractRunText(m[0]));
  }
  if (!pieces.length) {
    const textRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    while ((m = textRegex.exec(paragraphXml)) !== null) pieces.push(decodeXml(m[1]));
  }
  return pieces.join('');
}

function collectParagraphXml(xml) {
  const paragraphs = [];
  const re = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    paragraphs.push({
      start: m.index,
      end: m.index + m[0].length,
      xml: m[0],
      text: extractParagraphText(m[0]).replace(/\s+/g, ' ').trim()
    });
  }
  return paragraphs;
}

function looksLikeLetterDate(text) {
  const clean = String(text || '').replace(/\s+/g, ' ').trim();
  if (!clean || clean.length > 90) return false;
  return (
    /^(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$/i.test(clean) ||
    /^DATE\s+\\@\s+"[^"]+"\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$/i.test(clean) ||
    /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(clean) ||
    /^\d{4}-\d{1,2}-\d{1,2}$/.test(clean)
  );
}

function paragraphPropertiesFromXml(paragraphXml) {
  const match = paragraphXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/);
  return match ? match[0] : '';
}

function blankParagraphLike(paragraphXml) {
  const open = (paragraphXml.match(/^<w:p\b[^>]*>/) || ['<w:p>'])[0];
  return `${open}${paragraphPropertiesFromXml(paragraphXml)}<w:r><w:t></w:t></w:r></w:p>`;
}

function ensureBlankParagraphsAfterDate(xml, count, log) {
  const wanted = Number(count) || 0;
  if (wanted <= 0) return { xml, inserted: 0, dateText: null, reason: 'disabled' };
  const paragraphs = collectParagraphXml(xml);
  const maxSearch = Math.min(paragraphs.length, 40);
  for (let i = 0; i < maxSearch; i++) {
    const p = paragraphs[i];
    if (!looksLikeLetterDate(p.text)) continue;
    let existingBlank = 0;
    let next = i + 1;
    while (next < paragraphs.length && !paragraphs[next].text) {
      existingBlank++;
      next++;
    }
    const missing = Math.max(0, wanted - existingBlank);
    if (!missing) return { xml, inserted: 0, dateText: p.text, existingBlank };
    const blanks = Array.from({ length: missing }, () => blankParagraphLike(p.xml)).join('');
    const out = xml.slice(0, p.end) + blanks + xml.slice(p.end);
    const dlog = typeof log === 'function' ? log : (() => {});
    dlog(`[docx-presentation] Added ${missing} blank paragraph(s) after date "${p.text}" (${existingBlank} already present).`);
    return { xml: out, inserted: missing, dateText: p.text, existingBlank };
  }
  return { xml, inserted: 0, dateText: null, reason: 'date paragraph not found near top of document' };
}

function removeHighlightMarkup(xml) {
  let removedHighlight = 0;
  let removedHighlightPairs = 0;
  const out = String(xml || '')
    .replace(/<w:highlight\b[^>]*\/>/g, () => {
      removedHighlight++;
      return '';
    })
    .replace(/<w:highlight\b[^>]*>[\s\S]*?<\/w:highlight>/g, () => {
      removedHighlightPairs++;
      return '';
    });
  return { xml: out, removedHighlight: removedHighlight + removedHighlightPairs };
}

function applyFinalDocxPresentation(filePath, opts = {}) {
  const log = typeof opts.log === 'function' ? opts.log : (() => {});
  const zip = new PizZip(fs.readFileSync(filePath));
  const result = {
    blankLinesAfterDate: Number(opts.blankLinesAfterDate) || 0,
    dateSpacing: null,
    removedHighlights: 0,
    touchedParts: []
  };

  for (const name of DOCX_TEXT_TARGETS) {
    const file = zip.file(name);
    if (!file) continue;
    let xml = file.asText();
    let changed = false;

    if (opts.removeHighlights !== false) {
      const highlightResult = removeHighlightMarkup(xml);
      xml = highlightResult.xml;
      if (highlightResult.removedHighlight) {
        result.removedHighlights += highlightResult.removedHighlight;
        changed = true;
      }
    }

    if (name === 'word/document.xml' && result.blankLinesAfterDate > 0) {
      const spacingResult = ensureBlankParagraphsAfterDate(xml, result.blankLinesAfterDate, log);
      xml = spacingResult.xml;
      result.dateSpacing = {
        inserted: spacingResult.inserted,
        dateText: spacingResult.dateText,
        existingBlank: spacingResult.existingBlank || 0,
        reason: spacingResult.reason || null
      };
      if (spacingResult.inserted) changed = true;
    }

    if (changed) {
      zip.file(name, xml);
      result.touchedParts.push(name);
    }
  }

  if (result.removedHighlights) {
    log(`[docx-presentation] Removed ${result.removedHighlights} highlight tag(s) from generated letter.`);
  } else if (opts.removeHighlights !== false) {
    log('[docx-presentation] No highlight tags found to remove.');
  }
  if (result.dateSpacing && result.dateSpacing.reason) {
    log(`[docx-presentation] Date spacing not changed: ${result.dateSpacing.reason}`);
  }

  fs.writeFileSync(filePath, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  return result;
}

function applyTemplateMarkers(xml, markerEvaluator, log) {
  if (typeof markerEvaluator !== 'function') return xml;
  const dlog = typeof log === 'function' ? log : (() => {});
  const paragraphRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let out = '';
  let last = 0;
  const stack = [];
  let m;

  while ((m = paragraphRe.exec(xml)) !== null) {
    out += xml.slice(last, m.index);
    last = m.index + m[0].length;

    const paragraph = m[0];
    const text = extractParagraphText(paragraph).replace(/\s+/g, ' ').trim();
    const marker = markerEvaluator(text);
    const insideDeletedBlock = stack.some((entry) => entry === false);

    if (marker && marker.kind === 'start') {
      stack.push(!!marker.keep);
      dlog(`[template-rule] ${marker.keep ? 'keep' : 'delete'} marker block: ${marker.reason || marker.raw || text}`);
      continue;
    }
    if (marker && marker.kind === 'end') {
      if (stack.length) stack.pop();
      continue;
    }
    if (insideDeletedBlock) continue;

    out += paragraph;
  }

  out += xml.slice(last);
  if (stack.length) dlog(`[template-rule] warning: ${stack.length} conditional marker block(s) missing [[END]]`);
  return out;
}

function extractDocxMarkedTargets(filePath) {
  const content = fs.readFileSync(filePath);
  const zip = new PizZip(content);
  const seen = new Set();
  const all = [];

  for (const target of DOCX_TEXT_TARGETS) {
    const file = zip.file(target);
    if (!file) continue;
    const marked = extractMarkedTargetsFromXml(file.asText(), target);
    for (const item of marked) {
      const key = `${item.text}|${item.yellowHighlight}|${item.redText}`;
      if (seen.has(key)) continue;
      seen.add(key);
      all.push(item);
    }
  }
  return all;
}

function extractDocxContext(filePath) {
  return {
    text: extractDocxText(filePath),
    markedTargets: extractDocxMarkedTargets(filePath)
  };
}

// Word often splits a literal string across multiple <w:r>/<w:t> runs because
// of tracked-changes, spell-check, or mid-word formatting. For each target
// string we want to replace, find paragraphs whose concatenated text contains
// it and rebuild them so the string lives in one run.
function mergeSplitTargets(xml, targets) {
  if (!targets || !targets.length) return xml;

  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    const textNodes = [];
    const textRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    let m;
    while ((m = textRegex.exec(paragraph)) !== null) {
      textNodes.push(m[1]);
    }
    if (textNodes.length < 2) return paragraph;
    const decode = (s) => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
    const decodedNodes = textNodes.map(decode);
    const merged = textNodes.join('');
    const decoded = decodedNodes.join('');

    const hit = targets.some((t) => decoded.includes(t));
    if (!hit) return paragraph;

    // If a target already fits in a single <w:t>, leave the paragraph alone ?
    // merging would strip tabs/breaks between runs (RE:\t1730 Bedford -> RE:1730 Bedford).
    const singleHit = targets.some((t) => decodedNodes.some((n) => n.includes(t)));
    if (singleHit) return paragraph;

    let first = true;
    return paragraph.replace(textRegex, () => {
      if (first) {
        first = false;
        return `<w:t xml:space="preserve">${merged}</w:t>`;
      }
      return `<w:t xml:space="preserve"></w:t>`;
    });
  });
}

function addParagraphProperties(paragraph, props) {
  const wantsKeepNext = !!props.keepNext || /<w:keepNext\b/.test(paragraph);
  const wantsKeepLines = !!props.keepLines || /<w:keepLines\b/.test(paragraph);
  const inserts = [];
  if (wantsKeepNext) inserts.push('<w:keepNext/>');
  if (wantsKeepLines) inserts.push('<w:keepLines/>');
  if (!inserts.length) return paragraph;

  const insertXml = inserts.join('');
  const stripExistingKeep = (pPr) => pPr
    .replace(/<w:keepNext\b[^>]*(?:\/>|>[\s\S]*?<\/w:keepNext>)/g, '')
    .replace(/<w:keepLines\b[^>]*(?:\/>|>[\s\S]*?<\/w:keepLines>)/g, '');
  const insertIntoPPr = (pPr) => {
    const cleaned = stripExistingKeep(pPr);
    // WordprocessingML paragraph properties are order-sensitive. Insert after
    // pStyle so Word does not repair the document, and normalize any existing
    // keepNext/keepLines tags that older generated files may have misplaced.
    const pStyle = '<w:pStyle\\b[^>]*(?:/>|>[\\s\\S]*?</w:pStyle>)';
    const insertPoint = new RegExp(`(<w:pPr\\b[^>]*>\\s*(?:${pStyle}\\s*)?)`);
    return cleaned.replace(insertPoint, `$1${insertXml}`);
  };
  if (/<w:pPr\b/.test(paragraph)) {
    if (/<w:pPr\b[^>]*\/>/.test(paragraph)) {
      return paragraph.replace(/<w:pPr\b([^>]*)\/>/, `<w:pPr$1>${insertXml}</w:pPr>`);
    }
    return paragraph.replace(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/, insertIntoPPr);
  }
  return paragraph.replace(/(<w:p\b[^>]*>)/, `$1<w:pPr>${insertXml}</w:pPr>`);
}

function keepSignatureBlockTogether(xml, log) {
  const paragraphs = [];
  const paragraphRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let m;
  while ((m = paragraphRe.exec(xml)) !== null) {
    paragraphs.push({
      start: m.index,
      end: m.index + m[0].length,
      xml: m[0],
      text: extractParagraphText(m[0]).replace(/\s+/g, ' ').trim()
    });
  }
  if (!paragraphs.length) return xml;

  const startIndex = paragraphs.findIndex((p) => /^Sincerely yours,?$/i.test(p.text));
  if (startIndex === -1) return xml;

  let endIndex = -1;
  let nonEmpty = 0;
  for (let i = startIndex; i < paragraphs.length && i <= startIndex + 8; i++) {
    const text = paragraphs[i].text;
    if (text) nonEmpty++;
    if (/Metropolitan Realty/i.test(text) || (nonEmpty >= 3 && i > startIndex)) {
      endIndex = i;
      break;
    }
  }
  if (endIndex === -1) endIndex = Math.min(paragraphs.length - 1, startIndex + 3);

  const replacements = new Map();
  for (let i = startIndex; i <= endIndex; i++) {
    replacements.set(
      i,
      addParagraphProperties(paragraphs[i].xml, {
        keepNext: i < endIndex,
        keepLines: true
      })
    );
  }

  const dlog = typeof log === 'function' ? log : (() => {});
  dlog(`[layout-fix] Keeping signature block together across paragraphs ${startIndex + 1}-${endIndex + 1}`);

  let out = '';
  let last = 0;
  for (let i = 0; i < paragraphs.length; i++) {
    out += xml.slice(last, paragraphs[i].start);
    out += replacements.get(i) || paragraphs[i].xml;
    last = paragraphs[i].end;
  }
  out += xml.slice(last);
  return out;
}

function isProtectedEmptyParagraph(paragraph) {
  return /<w:br\b|<w:drawing\b|<w:pict\b|<w:object\b|<w:fldChar\b/i.test(paragraph);
}

function isEmptyParagraph(paragraph) {
  return !extractParagraphText(paragraph).replace(/\s+/g, '').trim() && !isProtectedEmptyParagraph(paragraph);
}

function isNumberedParagraph(paragraph) {
  return /<w:numPr\b/i.test(paragraph) || /<w:pStyle\b[^>]*w:val="[^"]*(?:List|Number|num|NoSpacing)[^"]*"/i.test(paragraph);
}

function tidyBlankParagraphs(xml, log) {
  const paragraphs = [];
  const paragraphRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let m;
  while ((m = paragraphRe.exec(xml)) !== null) {
    paragraphs.push({
      start: m.index,
      end: m.index + m[0].length,
      xml: m[0],
      empty: isEmptyParagraph(m[0]),
      numbered: isNumberedParagraph(m[0])
    });
  }
  if (!paragraphs.length) return xml;

  const remove = new Set();
  let removed = 0;
  for (let i = 0; i < paragraphs.length; i++) {
    if (!paragraphs[i].empty) continue;
    const prev = i > 0 ? paragraphs[i - 1] : null;
    if (prev && prev.empty && !remove.has(i - 1)) {
      remove.add(i);
      removed++;
    }
  }

  if (!remove.size) return xml;
  const dlog = typeof log === 'function' ? log : (() => {});
  dlog(`[layout-fix] Removed ${removed} extra blank paragraph(s)`);

  let out = '';
  let last = 0;
  for (let i = 0; i < paragraphs.length; i++) {
    out += xml.slice(last, paragraphs[i].start);
    if (!remove.has(i)) out += paragraphs[i].xml;
    last = paragraphs[i].end;
  }
  out += xml.slice(last);
  return out;
}

function removeDuplicateProjectedAssessedHeadings(xml, log) {
  const target = 'projected assessed value increase and phase-in';
  const paragraphs = [];
  const paragraphRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let m;
  while ((m = paragraphRe.exec(xml)) !== null) {
    paragraphs.push({
      start: m.index,
      end: m.index + m[0].length,
      xml: m[0],
      text: extractParagraphText(m[0]).replace(/\s+/g, ' ').trim().toLowerCase()
    });
  }

  let seen = false;
  const remove = new Set();
  for (let i = 0; i < paragraphs.length; i++) {
    if (paragraphs[i].text !== target) continue;
    if (!seen) {
      seen = true;
      continue;
    }
    remove.add(i);
  }
  if (!remove.size) return xml;

  const dlog = typeof log === 'function' ? log : (() => {});
  dlog(`[layout-fix] Removed ${remove.size} duplicate Projected Assessed Value Increase heading(s)`);

  let out = '';
  let last = 0;
  for (let i = 0; i < paragraphs.length; i++) {
    out += xml.slice(last, paragraphs[i].start);
    if (!remove.has(i)) out += paragraphs[i].xml;
    last = paragraphs[i].end;
  }
  out += xml.slice(last);
  return out;
}

// Phase 1: swap-mode filler. Takes a list of swaps [{ fieldName, oldValue, newValue }]
// and produces the output file. Returns { applied, missed } for the UI summary.
function fillDocxSwaps(templatePath, swaps, outputPath, opts) {
  opts = opts || {};
  const content = fs.readFileSync(templatePath);
  const zip = new PizZip(content);

  // Sort longest-first so "99 residential rental units" is replaced before "99".
  const sorted = [...swaps].sort((a, b) => b.oldValue.length - a.oldValue.length);
  const oldValues = sorted.map((s) => s.oldValue);

  const hitCounts = new Map(); // oldValue -> total replacements across all files

  for (const target of DOCX_TEXT_TARGETS) {
    const file = zip.file(target);
    if (!file) continue;
    let xml = file.asText();

    if (opts.markerEvaluator) {
      xml = applyTemplateMarkers(xml, opts.markerEvaluator, opts.log);
    }

    // Merge runs so split old values become contiguous text in one <w:t>.
    xml = mergeSplitTargets(xml, oldValues);

    for (const swap of sorted) {
      const oldEsc = escapeXml(swap.oldValue);
      let count = 0;
      if (swap.newValue === '') {
        let result = xml;
        while (true) {
          const next = removeParagraphContaining(result, oldEsc);
          if (next == null) break;
          result = next;
          count++;
        }
        xml = result;
      } else {
        const newEsc = escapeXml(swap.newValue);
        const re = new RegExp(escapeRegex(oldEsc), 'g');
        xml = xml.replace(re, () => { count++; return newEsc; });
      }
      if (count) {
        hitCounts.set(swap.oldValue, (hitCounts.get(swap.oldValue) || 0) + count);
      }
    }
    if (target === 'word/document.xml') {
      xml = removeDuplicateProjectedAssessedHeadings(xml, opts.log);
      xml = tidyBlankParagraphs(xml, opts.log);
      xml = keepSignatureBlockTogether(xml, opts.log);
    }
    if (/^word\/.*\.xml$/i.test(target)) {
      xml = normalizeParagraphPropertyOrder(xml, opts.log);
    }
    zip.file(target, xml);
  }

  const applied = [];
  const missed = [];
  for (const swap of sorted) {
    const count = hitCounts.get(swap.oldValue) || 0;
    if (count > 0) applied.push({ ...swap, count });
    else missed.push(swap);
  }

  const outputBuffer = zip.generate({ type: 'nodebuffer' });
  fs.writeFileSync(outputPath, outputBuffer);
  return { outputPath, applied, missed };
}

module.exports = {
  extractDocxText,
  extractDocxMarkedTargets,
  extractDocxContext,
  fillDocxSwaps,
  validateDocx,
  applyFinalDocxPresentation
};
