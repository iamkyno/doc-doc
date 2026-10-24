// Headers and footers in .docx files.
//
// DocDoc keeps a simple model: header and footer each have left / centre / right text,
// with tokens {page}, {pages}, {date} and {title}, plus "different first page".
//   readHeaderFooter(zip)            -> model (or null) from a Word file
//   writeHeaderFooter(zip, hf, page) -> adds Word header/footer parts with real page fields

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const CT_HEADER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const CT_FOOTER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

const empty = () => ({ l: '', c: '', r: '' });
const isEmpty = (s) => !s || !(s.l || s.c || s.r);

function decodeXml(s) {
  return s
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, '&');
}
function encodeXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function fieldToken(instr) {
  const name = (instr || '').trim().split(/\s+/)[0].toUpperCase();
  if (name === 'PAGE') return '{page}';
  if (name === 'NUMPAGES' || name === 'SECTIONPAGES') return '{pages}';
  if (name === 'DATE' || name === 'CREATEDATE' || name === 'SAVEDATE' || name === 'PRINTDATE') return '{date}';
  if (name === 'TITLE' || name === 'FILENAME') return '{title}';
  return null;
}

// Text of one paragraph, with tabs as \t and page fields as tokens.
function paragraphText(px) {
  let out = '';
  let field = null; // { instr, result, inResult }
  const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>|<w:fldSimple\b[^>]*w:instr="([^"]*)"[^>]*>([\s\S]*?)<\/w:fldSimple>|<w:fldChar\b[^>]*w:fldCharType="(begin|separate|end)"[^>]*\/>|<w:instrText(?:\s[^>]*)?>([^<]*)<\/w:instrText>/g;
  let m;
  while ((m = re.exec(px))) {
    if (m[0].startsWith('<w:fldSimple')) {
      const tok = fieldToken(decodeXml(m[2]));
      out += tok || decodeXml((m[3].match(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join(''));
    } else if (m[4] === 'begin') {
      field = { instr: '', result: '', inResult: false };
    } else if (m[4] === 'separate') {
      if (field) field.inResult = true;
    } else if (m[4] === 'end') {
      if (field) out += fieldToken(field.instr) || field.result;
      field = null;
    } else if (m[5] !== undefined) {
      if (field) field.instr += decodeXml(m[5]);
    } else if (m[1] !== undefined) {
      const t = decodeXml(m[1]);
      if (field) { if (field.inResult) field.result += t; } else out += t;
    } else if (m[0] === '<w:tab/>') {
      if (!field) out += '\t';
    } else if (m[0] === '<w:br/>') {
      if (!field) out += '\n';
    }
  }
  return out;
}

// One header/footer part -> { l, c, r }
function partToSlots(xml) {
  const slots = empty();
  const add = (k, t) => {
    t = t.replace(/\s+$/g, '').replace(/^\s+/g, '');
    if (!t) return;
    slots[k] = slots[k] ? slots[k] + '\n' + t : t;
  };
  const table = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/);
  if (table) {
    // a borderless table is a common way to lay out left / centre / right
    const row = (table[0].match(/<w:tr\b[\s\S]*?<\/w:tr>/) || [''])[0];
    const cells = (row.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []).map((c) => (c.match(/<w:p\b[\s\S]*?<\/w:p>/g) || []).map(paragraphText).join('\n'));
    if (cells.length >= 3) { add('l', cells[0]); add('c', cells.slice(1, -1).join(' ')); add('r', cells[cells.length - 1]); }
    else if (cells.length === 2) { add('l', cells[0]); add('r', cells[1]); }
    else if (cells.length === 1) add('l', cells[0]);
    xml = xml.replace(table[0], '');
  }
  for (const p of xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || []) {
    const text = paragraphText(p);
    if (!text.trim()) continue;
    const jc = (p.match(/<w:jc w:val="(\w+)"/) || [])[1] || 'left';
    if (text.includes('\t')) {
      const parts = text.split('\t');
      if (parts.length >= 3) { add('l', parts[0]); add('c', parts.slice(1, -1).join(' ')); add('r', parts[parts.length - 1]); }
      else if (jc === 'center') { add('c', parts.join(' ')); }
      else { add('l', parts[0]); add('r', parts[1]); }
    } else if (jc === 'center') add('c', text);
    else if (jc === 'right' || jc === 'end') add('r', text);
    else add('l', text);
  }
  return slots;
}

async function readHeaderFooter(zip) {
  const docFile = zip.file('word/document.xml');
  const relFile = zip.file('word/_rels/document.xml.rels');
  if (!docFile || !relFile) return null;
  const doc = await docFile.async('string');
  const rels = await relFile.async('string');
  // the section properties that apply to the end of the document (the main section)
  const sects = doc.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g);
  if (!sects) return null;
  const sect = sects[sects.length - 1];
  const target = (id) => {
    const m = rels.match(new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*>`));
    const t = m && (m[0].match(/Target="([^"]+)"/) || [])[1];
    return t ? 'word/' + t.replace(/^\/?word\//, '').replace(/^\.\//, '') : null;
  };
  const ref = (kind, type) => {
    const m = sect.match(new RegExp(`<w:${kind}Reference\\b[^>]*w:type="${type}"[^>]*\\/>`));
    return m && (m[0].match(/r:id="([^"]+)"/) || [])[1];
  };
  const load = async (kind) => {
    const id = ref(kind, 'default');
    const f = id && target(id) && zip.file(target(id));
    return f ? partToSlots(await f.async('string')) : empty();
  };
  const header = await load('header');
  const footer = await load('footer');
  const firstDifferent = /<w:titlePg(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(sect);
  if (isEmpty(header) && isEmpty(footer)) return null;
  return { header, footer, firstDifferent };
}

/* ---------- writing ---------- */

function runs(text, title) {
  const rPr = '<w:rPr><w:sz w:val="18"/><w:color w:val="444444"/></w:rPr>';
  const field = (instr, shown) => `<w:fldSimple w:instr=" ${encodeXml(instr)} "><w:r>${rPr}<w:t>${shown}</w:t></w:r></w:fldSimple>`;
  let out = '';
  const parts = String(text || '').split(/(\{page\}|\{pages\}|\{date\}|\{title\}|\n)/);
  for (const part of parts) {
    if (!part) continue;
    if (part === '{page}') out += field('PAGE', '1');
    else if (part === '{pages}') out += field('NUMPAGES', '1');
    else if (part === '{date}') out += field('DATE \\@ "d MMMM yyyy"', '');
    else if (part === '{title}') out += `<w:r>${rPr}<w:t xml:space="preserve">${encodeXml(title || '')}</w:t></w:r>`;
    else if (part === '\n') out += `<w:r>${rPr}<w:br/></w:r>`;
    else out += `<w:r>${rPr}<w:t xml:space="preserve">${encodeXml(part)}</w:t></w:r>`;
  }
  return out;
}

// left <tab> centre <tab> right, using a centre and a right tab stop
function partXml(kind, slots, widthTwips, title) {
  const tag = kind === 'header' ? 'w:hdr' : 'w:ftr';
  const tab = `<w:r><w:tab/></w:r>`;
  const body = slots
    ? runs(slots.l, title) + tab + runs(slots.c, title) + tab + runs(slots.r, title)
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${tag} xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:p><w:pPr><w:tabs><w:tab w:val="center" w:pos="${Math.round(widthTwips / 2)}"/><w:tab w:val="right" w:pos="${widthTwips}"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${body}</w:p></${tag}>`;
}

async function writeHeaderFooter(zip, hf, page, title) {
  if (!hf || (isEmpty(hf.header) && isEmpty(hf.footer))) return;
  const twip = (mm) => Math.round(mm * 56.6929);
  const pageW = page.landscape ? page.heightMm : page.widthMm;
  const width = twip(pageW - 2 * page.marginMm);

  const files = [
    { name: 'docdocHeader1.xml', kind: 'header', type: 'default', slots: hf.header },
    { name: 'docdocFooter1.xml', kind: 'footer', type: 'default', slots: hf.footer },
  ];
  if (hf.firstDifferent) {
    files.push({ name: 'docdocHeader2.xml', kind: 'header', type: 'first', slots: null });
    files.push({ name: 'docdocFooter2.xml', kind: 'footer', type: 'first', slots: null });
  }

  let rels = await zip.file('word/_rels/document.xml.rels').async('string');
  let ct = await zip.file('[Content_Types].xml').async('string');
  let doc = await zip.file('word/document.xml').async('string');
  let refs = '';
  files.forEach((f, i) => {
    const id = `rIdDocDocHF${i + 1}`;
    zip.file(`word/${f.name}`, partXml(f.kind, f.slots, width, title));
    rels = rels.replace('</Relationships>', `<Relationship Id="${id}" Type="${f.kind === 'header' ? REL_HEADER : REL_FOOTER}" Target="${f.name}"/></Relationships>`);
    ct = ct.replace('</Types>', `<Override PartName="/word/${f.name}" ContentType="${f.kind === 'header' ? CT_HEADER : CT_FOOTER}"/></Types>`);
    refs += `<w:${f.kind}Reference w:type="${f.type}" r:id="${id}"/>`;
  });
  // header/footer references must come first inside the section properties
  doc = doc.replace(/<w:sectPr\b[^>]*>(?![\s\S]*<w:sectPr\b)/, (m) => m + refs);
  // distance of header/footer from the page edge: half the margin (like Word's 1.25 cm)
  const dist = Math.round(twip(page.marginMm) / 2);
  doc = doc.replace(/(<w:pgMar\b[^>]*?)\sw:header="\d+"/, '$1').replace(/(<w:pgMar\b[^>]*?)\sw:footer="\d+"/, '$1');
  doc = doc.replace(/<w:pgMar\b/, `<w:pgMar w:header="${dist}" w:footer="${dist}"`);
  if (hf.firstDifferent && !/<w:titlePg/.test(doc)) doc = doc.replace(/<\/w:sectPr>(?![\s\S]*<\/w:sectPr>)/, '<w:titlePg/></w:sectPr>');
  if (!/xmlns:r=/.test(doc.slice(0, 2000))) doc = doc.replace(/<w:document\b/, `<w:document xmlns:r="${R_NS}"`);
  zip.file('word/_rels/document.xml.rels', rels);
  zip.file('[Content_Types].xml', ct);
  zip.file('word/document.xml', doc);
}

module.exports = { readHeaderFooter, writeHeaderFooter, partToSlots };
