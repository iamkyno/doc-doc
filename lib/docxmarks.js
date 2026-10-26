// Word details the HTML converter can't write.
//
// Before conversion the app puts invisible marker characters into the HTML (see MK in
// src/app.js). This turns them into real Word markup in word/document.xml:
//   paragraph:  h=567;f=360;t=dot:8640   -> hanging / first-line indent, tab stops
//   run text:   -> <w:tab/>,   -> small caps,   -> all caps
//   before a picture:  /  -> text wraps around it (right / left),
//                      n  -> lower it by n half-points (equations)

const MARKS = /[-]/;

// Child element order Word expects inside <w:pPr> and <w:rPr>.
const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
  'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
  'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
  'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
  'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern',
  'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em',
  'lang', 'eastAsianLayout', 'specVanish', 'oMath'];

// Top-level child elements of a property block's inner XML: [{ name, start }]
function children(inner) {
  const out = [];
  let depth = 0;
  const re = /<(\/?)w:(\w+)\b[^>]*?(\/?)>/g;
  let m;
  while ((m = re.exec(inner))) {
    if (m[1]) { depth--; continue; }
    if (depth === 0) out.push({ name: m[2], start: m.index });
    if (!m[3]) depth++;
  }
  return out;
}
function insertOrdered(inner, name, xml, order) {
  const rank = order.indexOf(name);
  const next = children(inner).find((c) => order.indexOf(c.name) > rank);
  return next ? inner.slice(0, next.start) + xml + inner.slice(next.start) : inner + xml;
}
// Replace (or add) the property block <w:tag>…</w:tag> at the start of a paragraph/run body.
function editProps(body, tag, fn) {
  const re = new RegExp(`<w:${tag}/>|<w:${tag}>([\\s\\S]*?)</w:${tag}>`);
  const m = body.match(re);
  if (m) return body.replace(m[0], `<w:${tag}>${fn(m[1] || '')}</w:${tag}>`);
  return `<w:${tag}>${fn('')}</w:${tag}>` + body;
}

let anchorId = 0;
// An inline picture that text wraps around, kept at the paragraph's left or right edge.
function toAnchor(inline, side) {
  const ext = (inline.match(/<wp:extent\b[^>]*\/>/) || [''])[0];
  const body = inline
    .replace(/^<wp:inline\b[^>]*>/, '').replace(/<\/wp:inline>$/, '')
    .replace(/<wp:extent\b[^>]*\/>/, '').replace(/<wp:effectExtent\b[^>]*\/>/, '');
  anchorId++;
  return `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251658240 + anchorId}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
    + '<wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="column"><wp:align>${side}</wp:align></wp:positionH>`
    + '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
    + `${ext}<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>${body}</wp:anchor>`;
}

function fixRun(open, body) {
  const flags = new Set();
  body = body.replace(/<w:t(\s[^>]*)?>([^<]*)<\/w:t>/g, (m, attrs, text) => {
    text = text.replace(/[]/g, (c) => { flags.add(c); return ''; });
    if (!text.includes('')) return `<w:t xml:space="preserve">${text}</w:t>`;
    return text.split('').map((t) => (t ? `<w:t xml:space="preserve">${t}</w:t>` : '')).join('<w:tab/>');
  });
  if (flags.size) {
    body = editProps(body, 'rPr', (inner) => {
      if (flags.has('') && !/<w:caps\b/.test(inner)) inner = insertOrdered(inner, 'caps', '<w:caps/>', RPR_ORDER);
      if (flags.has('') && !/<w:smallCaps\b/.test(inner)) inner = insertOrdered(inner, 'smallCaps', '<w:smallCaps/>', RPR_ORDER);
      return inner;
    });
  }
  return `${open}${body}</w:r>`;
}

function fixParagraph(p) {
  if (!MARKS.test(p)) return p;
  // paragraph settings
  const props = [];
  p = p.replace(/([^<]*)/g, (m, s) => { props.push(...s.split(';').filter(Boolean)); return ''; });
  // pictures: wrapping and lowering apply to the next picture after the marker
  for (;;) {
    const m = /[]|(\d+)/.exec(p);
    if (!m) break;
    p = p.slice(0, m.index) + p.slice(m.index + m[0].length);
    const j = p.indexOf('<wp:inline', m.index);
    if (j < 0) continue;
    if (m[1] !== undefined) {
      // lower the run that holds the picture
      const rs = p.lastIndexOf('<w:r>', j);
      const re = p.indexOf('</w:r>', j);
      if (rs < 0 || re < 0) continue;
      const run = p.slice(rs + 5, re);
      const fixed = editProps(run, 'rPr', (inner) => (/<w:position\b/.test(inner) ? inner
        : insertOrdered(inner, 'position', `<w:position w:val="-${m[1]}"/>`, RPR_ORDER)));
      p = p.slice(0, rs + 5) + fixed + p.slice(re);
    } else {
      const k = p.indexOf('</wp:inline>', j);
      if (k < 0) continue;
      p = p.slice(0, j) + toAnchor(p.slice(j, k + 12), m[0] === '' ? 'left' : 'right') + p.slice(k + 12);
    }
  }
  // runs: tabs, small caps, caps
  p = p.replace(/(<w:r(?:\s[^>]*)?>)([\s\S]*?)<\/w:r>/g, (m, open, body) => fixRun(open, body));
  if (props.length) {
    let hanging = 0;
    let first = 0;
    const tabs = [];
    for (const pr of props) {
      const [k, v] = pr.split('=');
      if (k === 'h') hanging = Math.round(+v) || 0;
      else if (k === 'f') first = Math.round(+v) || 0;
      else if (k === 't') {
        const [leader, pos] = v.split(':');
        if (+pos > 0 && /^(none|dot|underscore|hyphen|heavy|middleDot)$/.test(leader)) tabs.push({ leader, pos: Math.round(+pos) });
      }
    }
    const open = p.match(/^<w:p(?:\s[^>]*)?>/)[0];
    let body = p.slice(open.length);
    body = editProps(body, 'pPr', (inner) => {
      if (tabs.length) {
        tabs.sort((a, b) => a.pos - b.pos);
        const xml = '<w:tabs>' + tabs.map((t) => `<w:tab w:val="left" w:leader="${t.leader}" w:pos="${t.pos}"/>`).join('') + '</w:tabs>';
        inner = /<w:tabs>/.test(inner) ? inner.replace(/<w:tabs>[\s\S]*?<\/w:tabs>/, xml) : insertOrdered(inner, 'tabs', xml, PPR_ORDER);
      }
      if (hanging || first) {
        const attr = hanging ? `w:hanging="${hanging}"` : `w:firstLine="${first}"`;
        const ind = inner.match(/<w:ind\b([^>]*)\/>/);
        if (ind) {
          const rest = ind[1].replace(/\s*w:(hanging|firstLine)="[^"]*"/g, '');
          inner = inner.replace(ind[0], `<w:ind${rest} ${attr}/>`);
        } else inner = insertOrdered(inner, 'ind', `<w:ind w:left="0" ${attr}/>`, PPR_ORDER);
      }
      return inner;
    });
    p = open + body;
  }
  return p;
}

function applyMarks(xml) {
  if (!MARKS.test(xml)) return xml;
  xml = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, fixParagraph);
  return xml.replace(/[-]/g, '');
}

// The converter writes the page settings (<w:sectPr>) as the body's first child; Word
// expects them last.
function sectPrLast(xml) {
  const m = xml.match(/<w:body>(\s*)(<w:sectPr\b[\s\S]*?<\/w:sectPr>)/);
  if (!m) return xml;
  xml = xml.replace(m[0], '<w:body>' + m[1]);
  return xml.replace(/<\/w:body>(?![\s\S]*<\/w:body>)/, m[2] + '</w:body>');
}

module.exports = { applyMarks, sectPrLast, insertOrdered, PPR_ORDER, RPR_ORDER };
