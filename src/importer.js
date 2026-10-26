/* High-fidelity .docx import.
 *
 * docx-preview renders the Word file with its real styles (themes, paragraph
 * and character styles, numbering, table styles). We then read the final
 * computed look of every element, write it inline, and rebuild the result as
 * plain editable HTML:
 *   - simple numbered / bulleted lists become real <ol>/<ul> (with the right start number)
 *   - outline numbering like 1.1 / 1.1.1 / (a) / (i) keeps Word's exact numbers as text
 *   - Word text boxes become DocDoc text boxes; floating pictures float
 *   - tab leaders ("Name: ........"), hidden text and merged table cells are respected
 */
(() => {
  'use strict';

  const PX_TO_MM = 25.4 / 96;
  const PT_TO_PX = 96 / 72;
  const MATHML = 'http://www.w3.org/1998/Math/MathML';
  const SVGNS = 'http://www.w3.org/2000/svg';
  const INHERITED = ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'color', 'letterSpacing', 'textTransform', 'fontVariantCaps'];
  const TRANSPARENT = (v) => !v || v === 'transparent' || v === 'rgba(0, 0, 0, 0)';
  const BLOCKS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'PRE']);

  function firstFamily(f) {
    return (f || '').split(',')[0].replace(/["']/g, '').trim();
  }
  function px(v) {
    return parseFloat(v) || 0;
  }
  function round(n) {
    return Math.round(n * 100) / 100;
  }
  // "Arial" -> "Arial, sans-serif", so a missing font falls back to something similar
  function withFallback(ff) {
    if (!ff || /\b(serif|sans-serif|monospace|cursive|fantasy|system-ui)\s*$/i.test(ff)) return ff;
    const name = firstFamily(ff).toLowerCase();
    const generic = /mono|courier|consol|lucida console/.test(name) ? 'monospace'
      : /times|georgia|cambria|caladea|garamond|book|roman|palatino|baskerville|bodoni|century|minion|serif/.test(name) ? 'serif'
        : 'sans-serif';
    return `${ff}, ${generic}`;
  }
  function parseStyleAttr(s) {
    const out = {};
    for (const part of String(s || '').split(';')) {
      const i = part.indexOf(':');
      if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
    }
    return out;
  }
  const ptToPx = (v) => (/pt$/.test(v) ? px(v) * PT_TO_PX : /in$/.test(v) ? px(v) * 96 : /cm$/.test(v) ? px(v) * 96 / 2.54 : px(v));

  // Read what an element looks like, as a plain object of CSS properties.
  function snapshot(el, parentCs, isBlock) {
    const cs = getComputedStyle(el);
    const out = {};
    for (const p of INHERITED) {
      if (!parentCs || cs[p] !== parentCs[p]) out[p] = cs[p];
    }
    const tag = el.tagName;
    if (!TRANSPARENT(cs.backgroundColor)) out.backgroundColor = cs.backgroundColor;
    if (cs.textDecorationLine && cs.textDecorationLine !== 'none') {
      out.textDecorationLine = cs.textDecorationLine;
      if (cs.textDecorationStyle && cs.textDecorationStyle !== 'solid') out.textDecorationStyle = cs.textDecorationStyle;
    }
    if (cs.verticalAlign === 'super' || cs.verticalAlign === 'sub' || (/px$/.test(cs.verticalAlign) && px(cs.verticalAlign))) {
      out.verticalAlign = cs.verticalAlign;
    }
    if (tag === 'SPAN') {
      if (cs.display === 'inline-block' && px(cs.width)) {
        out.display = 'inline-block';
        out.width = cs.width;
      }
      if (cs.wordSpacing && px(cs.wordSpacing)) out.wordSpacing = cs.wordSpacing; // tab stops
    }
    if (isBlock) {
      out.textAlign = cs.textAlign === 'start' ? 'left' : cs.textAlign;
      if (cs.textIndent && px(cs.textIndent)) out.textIndent = cs.textIndent;
      const lh = cs.lineHeight;
      if (lh && lh !== 'normal') out.lineHeight = String(round(px(lh) / px(cs.fontSize)));
      for (const side of ['Top', 'Bottom', 'Left', 'Right']) {
        const m = cs['margin' + side];
        if (px(m)) out['margin' + side] = m;
        const pd = cs['padding' + side];
        if (px(pd)) out['padding' + side] = pd;
        if (cs['border' + side + 'Style'] !== 'none' && px(cs['border' + side + 'Width'])) {
          out['border' + side] = `${cs['border' + side + 'Width']} ${cs['border' + side + 'Style']} ${cs['border' + side + 'Color']}`;
        }
      }
    }
    if (tag === 'TD' || tag === 'TH') {
      for (const side of ['Top', 'Bottom', 'Left', 'Right']) {
        const w = cs['border' + side + 'Width'];
        out['border' + side] = cs['border' + side + 'Style'] !== 'none' && px(w)
          ? `${w} ${cs['border' + side + 'Style']} ${cs['border' + side + 'Color']}`
          : 'none';
        out['padding' + side] = cs['padding' + side];
      }
      out.verticalAlign = cs.verticalAlign;
      if (el.style.width) out.width = cs.width;
    }
    if (tag === 'TR' && el.style.height) out.height = cs.height;
    if (tag === 'TABLE') {
      out.borderCollapse = cs.borderCollapse;
      if (el.style.width && el.style.width !== 'auto') out.width = cs.width;
      if (px(cs.marginLeft)) out.marginLeft = cs.marginLeft;
    }
    return { cs, out };
  }

  function applyStyle(el, obj) {
    el.removeAttribute('style');
    // font sizes in points, like Word
    if (obj.fontSize && /px$/.test(obj.fontSize)) obj.fontSize = Math.round(px(obj.fontSize) * 0.75 * 2) / 2 + 'pt';
    if (obj.fontFamily) obj.fontFamily = withFallback(obj.fontFamily);
    for (const [k, v] of Object.entries(obj)) el.style[k] = v;
    if (!el.getAttribute('style')) el.removeAttribute('style');
  }

  // If every text run in a block shares a font/size/colour, move it up to the block.
  const HOIST = ['fontFamily', 'fontSize', 'color', 'fontWeight', 'fontStyle'];
  function hoist(block, styles) {
    const runs = [];
    const walk = (n) => {
      for (const c of n.childNodes) {
        if (c.nodeType === 3 && c.nodeValue.trim()) runs.push(c.parentElement);
        else if (c.nodeType === 1 && !BLOCKS.has(c.tagName) && c.tagName !== 'TABLE' && c.namespaceURI !== SVGNS && c.namespaceURI !== MATHML) walk(c);
      }
    };
    walk(block);
    const own = styles.get(block);
    if (!own) return;
    if (!runs.length) {
      HOIST.forEach((p) => delete own[p]);
      return;
    }
    for (const p of HOIST) {
      // use the most common value among the runs (weighted by text length)
      const count = new Map();
      for (const r of runs) {
        const v = getComputedStyle(r)[p];
        count.set(v, (count.get(v) || 0) + r.textContent.length);
      }
      let v = null;
      let best = -1;
      for (const [k, n] of count) if (n > best) { best = n; v = k; }
      if (v == null) continue;
      own[p] = v;
      // children no longer need it
      block.querySelectorAll('*').forEach((d) => {
        const s = styles.get(d);
        if (!s) return;
        if (s[p] === v) delete s[p];
        // runs that differ from the new paragraph value must state their own
        else if (!(p in s) && d.firstChild && getComputedStyle(d)[p] !== v && d.parentElement === block) s[p] = getComputedStyle(d)[p];
      });
    }
  }

  function headingTag(cls) {
    const m = /(?:^|_)(heading([1-6])|title|subtitle)\b/i.exec(cls || '');
    if (!m) return null;
    if (/title/i.test(m[1]) && !/subtitle/i.test(m[1])) return 'H4';
    if (m[2]) return +m[2] <= 3 ? 'H' + m[2] : null;
    return null;
  }

  /* ---------- numbering ---------- */
  // docx-preview draws Word numbers with CSS counters. We replay those counters in
  // document order to know the exact number of every paragraph.
  function counterList(s, dflt) {
    const out = [];
    if (!s || s === 'none') return out;
    const parts = s.trim().split(/\s+/);
    for (let i = 0; i < parts.length; i++) {
      const name = parts[i];
      let v = dflt;
      if (i + 1 < parts.length && /^-?\d+$/.test(parts[i + 1])) v = +parts[++i];
      out.push([name, v]);
    }
    return out;
  }
  function cssUnescape(s) {
    return s.replace(/\\([0-9a-fA-F]{1,6})\s?/g, (m, h) => String.fromCodePoint(parseInt(h, 16))).replace(/\\(.)/g, '$1');
  }
  function contentTokens(content) {
    const out = [];
    const re = /"((?:[^"\\]|\\.)*)"|counter\(\s*([^,\s)]+)\s*(?:,\s*([a-z-]+))?\s*\)/g;
    let m;
    while ((m = re.exec(content || ''))) {
      if (m[2]) out.push({ counter: m[2], style: m[3] || 'decimal' });
      else out.push({ text: cssUnescape(m[1]) });
    }
    return out;
  }
  function toAlpha(n, upper) {
    let s = '';
    while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); }
    return upper ? s.toUpperCase() : s;
  }
  function toRoman(n, upper) {
    if (n <= 0) return String(n);
    const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let s = '';
    for (const [v, r] of map) while (n >= v) { s += r; n -= v; }
    return upper ? s.toUpperCase() : s;
  }
  function formatNumber(n, style) {
    switch (style) {
      case 'lower-alpha': case 'lower-latin': return toAlpha(n, false);
      case 'upper-alpha': case 'upper-latin': return toAlpha(n, true);
      case 'lower-roman': return toRoman(n, false);
      case 'upper-roman': return toRoman(n, true);
      case 'decimal-leading-zero': return (n < 10 ? '0' : '') + n;
      case 'none': return '';
      default: return String(n);
    }
  }
  const LIST_TYPES = new Set(['decimal', 'lower-alpha', 'upper-alpha', 'lower-roman', 'upper-roman', 'lower-latin', 'upper-latin']);
  function bulletType(text) {
    const c = (text || '').trim();
    if (!c || /^[\uf0b7•●\u2022\uf06c]/.test(c)) return 'disc';
    if (/^[o◦\u25e6]/.test(c)) return 'circle';
    if (/^[\uf0a7\uf0d8\uf0fc▪■§\u25aa]/.test(c)) return 'square';
    if (/^[\uf000-\uf8ff]/.test(c)) return 'disc'; // symbol-font bullet with no Unicode equivalent
    return `"${c.replace(/"/g, '')} "`;
  }

  function computeNumbering(host) {
    const counters = new Map();
    // docx-preview puts the starting values on :root
    for (const [n, v] of counterList(getComputedStyle(document.documentElement).counterReset, 0)) {
      if (/ddimp-num-/.test(n)) counters.set(n, v);
    }
    const info = new Map();
    const byNum = new Map();
    for (const p of host.querySelectorAll('p')) {
      const m = /(?:^|\s)ddimp-num-(\d+)-(\d+)(?:\s|$)/.exec(p.className);
      if (!m) continue;
      const cs = getComputedStyle(p);
      const bcs = getComputedStyle(p, '::before');
      for (const [n, v] of counterList(cs.counterReset, 0)) counters.set(n, v);
      for (const [n, v] of counterList(cs.counterSet, 0)) counters.set(n, v);
      for (const [n, v] of counterList(bcs.counterIncrement, 1)) counters.set(n, (counters.get(n) || 0) + v);
      const tokens = contentTokens(bcs.content);
      const own = `ddimp-num-${m[1]}-${m[2]}`;
      const label = tokens.map((t) => (t.counter ? formatNumber(counters.get(t.counter) || 0, t.style) : t.text)).join('').replace(/[\t\s]+$/, '');
      const hasCounter = tokens.some((t) => t.counter);
      const pattern = tokens.map((t) => (t.counter ? (t.counter === own ? '#' : '@') : t.text)).join('').replace(/[\t\s]+$/, '').replace(/^\s+/, '');
      const ownTok = tokens.find((t) => t.counter === own);
      const rec = {
        numId: m[1],
        level: +m[2],
        label,
        bullet: !hasCounter,
        bulletText: !hasCounter ? label : '',
        simple: hasCounter && pattern === '#.' && ownTok && LIST_TYPES.has(ownTok.style),
        type: ownTok ? ownTok.style : 'decimal',
        value: counters.get(own) || 1,
        heading: !!headingTag(p.className),
        labelCss: { fontWeight: bcs.fontWeight, fontStyle: bcs.fontStyle, color: bcs.color, textDecorationLine: bcs.textDecorationLine },
        pCss: { fontWeight: cs.fontWeight, fontStyle: cs.fontStyle, color: cs.color },
      };
      info.set(p, rec);
      if (!byNum.has(rec.numId)) byNum.set(rec.numId, []);
      byNum.get(rec.numId).push(rec);
    }
    // a list keeps real <ol>/<ul> only if every item of it is a plain "1." / "a." style or a bullet
    for (const recs of byNum.values()) {
      const real = recs.every((r) => !r.heading && (r.bullet || r.simple));
      for (const r of recs) r.mode = real ? 'list' : 'label';
    }
    // where the paragraph text starts after the number, to line text up like Word
    for (const [p, r] of info) {
      if (!r.label) continue;
      const cs = getComputedStyle(p);
      const hanging = -px(cs.textIndent);
      // with a hanging indent the number's tab reaches the indent; otherwise the next 1.27 cm tab stop
      r.labelWidth = hanging > 4 ? hanging : 48;
    }
    return info;
  }

  /* ---------- Word text boxes and shapes ---------- */
  // docx-preview draws them from Word's VML fallback, but without the box's line,
  // fill and wrapping. We read those from the file, in document order.
  async function readVmlBoxes(bytes) {
    try {
      const zip = await window.JSZip.loadAsync(bytes);
      const xml = await zip.file('word/document.xml').async('string');
      const out = [];
      const re = /<v:(shape|rect|roundrect|oval)\b([^>]*)>([\s\S]*?)<\/v:\1>/g;
      let m;
      while ((m = re.exec(xml))) {
        const attrs = m[2];
        const inner = m[3];
        const a = (name) => (attrs.match(new RegExp(`\\b${name}="([^"]*)"`)) || [])[1];
        out.push({
          textbox: /<v:textbox\b/.test(inner),
          filled: !/^(f|false)$/i.test(a('filled') || ''),
          fill: a('fillcolor') || (/<v:fill\b[^>]*color="([^"]+)"/.exec(inner) || [])[1] || '#ffffff',
          stroked: !/^(f|false)$/i.test(a('stroked') || ''),
          stroke: a('strokecolor') || '#000000',
          weight: a('strokeweight') || '0.75pt',
          dash: (/<v:stroke\b[^>]*dashstyle="(\w+)"/.exec(inner) || [])[1] || '',
          wrap: (/<w10:wrap\b[^>]*type="(\w+)"/.exec(inner) || [])[1] || 'none',
          round: m[1] === 'roundrect',
        });
      }
      return out;
    } catch (_) {
      return [];
    }
  }
  function vmlColor(c) {
    c = (c || '').trim().split(/\s/)[0];
    if (/^#[0-9a-f]{3,6}$/i.test(c)) return c;
    const named = { black: '#000000', white: '#ffffff', red: '#ff0000', blue: '#0000ff', green: '#008000', yellow: '#ffff00', gray: '#808080', silver: '#c0c0c0' };
    return named[c.toLowerCase()] || '#000000';
  }

  function makeBox(floating) {
    const box = document.createElement('div');
    box.className = floating ? 'textbox floating' : 'textbox';
    box.contentEditable = 'false';
    const body = document.createElement('div');
    body.className = 'tb-body';
    body.contentEditable = 'true';
    box.appendChild(body);
    return { box, body };
  }

  async function importDocx(bytes) {
    const host = document.createElement('div');
    // laid out in place but invisible: docx-preview measures positions (tab stops) on screen
    host.style.cssText = 'position:fixed;left:0;top:0;visibility:hidden;pointer-events:none;z-index:-10;contain:paint';
    const styleHost = document.createElement('div');
    document.body.append(host, styleHost);
    try {
      await window.docx.renderAsync(bytes, host, styleHost, {
        className: 'ddimp',
        inWrapper: false,
        breakPages: true,
        ignoreLastRenderedPageBreak: true,
        ignoreWidth: false,
        ignoreHeight: false,
        renderHeaders: false,
        renderFooters: false,
        renderFootnotes: true,
        renderEndnotes: true,
        renderComments: false,
        renderChanges: false,
        useBase64URL: true,
        experimental: true,
      });
      // docx-preview sizes tab stops (and draws their dotted / underlined leaders) half a second later
      if (host.querySelector('.ddimp-tab-stop')) await new Promise((r) => setTimeout(r, 650));
      const vml = await readVmlBoxes(bytes);

      const sections = Array.from(host.querySelectorAll('section.ddimp'));
      if (!sections.length) throw new Error('Nothing to show in this document');

      // Page set-up from the first section.
      const s0 = getComputedStyle(sections[0]);
      const wMm = px(sections[0].style.width ? s0.width : 0) * PX_TO_MM;
      const hMm = px(sections[0].style.minHeight ? s0.minHeight : 0) * PX_TO_MM;
      const marginPx = px(s0.paddingLeft);
      const contentW = px(s0.width) - px(s0.paddingLeft) - px(s0.paddingRight);
      const page = {
        widthMm: wMm,
        heightMm: hMm,
        marginMm: round(marginPx * PX_TO_MM) || 25.4,
      };

      const numbering = computeNumbering(host);

      // 1) Snapshot the final look of everything while the original styles still apply.
      const styles = new Map();
      const meta = new Map();
      const hidden = [];
      const drawings = []; // VML text boxes / shapes, in document order
      for (const sec of sections) {
        const article = sec.querySelector(':scope > article') || sec;
        const walk = (el, parentCs) => {
          if (el.namespaceURI === MATHML) return; // equations stay as they are
          if (el.namespaceURI === SVGNS) {
            if (el.tagName.toLowerCase() !== 'svg') return;
            const isBox = !!el.querySelector('foreignObject p');
            drawings.push({ svg: el, isBox, style: parseStyleAttr(el.getAttribute('style')), w: el.getBoundingClientRect().width, h: el.getBoundingClientRect().height });
            if (isBox) for (const fo of el.querySelectorAll('foreignObject')) for (const c of fo.children) walk(c, null);
            return;
          }
          const cs0 = getComputedStyle(el);
          if (cs0.display === 'none') { hidden.push(el); return; }
          const isBlock = BLOCKS.has(el.tagName) || el.tagName === 'DIV' && cs0.display === 'block';
          const { cs, out } = snapshot(el, parentCs, isBlock);
          styles.set(el, out);
          if (el.tagName === 'P') meta.set(el, { heading: headingTag(el.className) });
          if (el.tagName === 'SPAN' && el.classList.contains('ddimp-tab-stop')) meta.set(el, { tab: true });
          if (el.tagName === 'IMG') {
            const r = el.getBoundingClientRect();
            meta.set(el, { w: Math.round(r.width), h: Math.round(r.height) });
          }
          if (el.tagName === 'DIV' && el.querySelector(':scope > img')) {
            meta.set(el, {
              wrapper: true,
              float: cs.float,
              block: el.style.display === 'block',
              wrapNone: el.style.position === 'relative' && el.style.width === '0px',
              align: el.style.textAlign || '',
              left: ptToPx(el.style.left || '0'),
              top: ptToPx(el.style.top || '0'),
            });
          }
          for (const c of el.children) walk(c, cs);
        };
        // Top-level blocks are compared against nothing, so they keep their full look.
        for (const c of article.children) walk(c, null);
        // footnotes / endnotes sit beside the article
        if (article !== sec) for (const c of sec.children) if (c !== article && c.tagName === 'OL') walk(c, null);
      }
      hidden.forEach((h) => h.remove()); // e.g. the hidden halves of merged table cells
      // Paragraph-level fonts: lift shared run formatting to the paragraph.
      for (const [el] of meta) if (el.tagName === 'P') hoist(el, styles);
      // Containers take their fonts from the editor; paragraphs carry their own.
      const DEFAULTS = { letterSpacing: 'normal', textTransform: 'none', fontVariantCaps: 'normal', fontWeight: '400', fontStyle: 'normal' };
      for (const [el, s] of styles) {
        for (const [k, v] of Object.entries(DEFAULTS)) if (s[k] === v) delete s[k];
        if (/^(TABLE|TBODY|THEAD|TR|TD|TH|DIV|OL|UL)$/.test(el.tagName)) HOIST.forEach((p) => delete s[p]);
        if (s.textAlign === 'left') delete s.textAlign;
      }

      // 2) Rebuild lists (only "real" lists; outline numbers become text labels).
      const listOf = (el) => {
        const r = el.tagName === 'P' ? numbering.get(el) : null;
        return r && r.mode === 'list' ? r : null;
      };
      const out = document.createElement('div');
      const convertSection = (article) => {
        const frag = document.createDocumentFragment();
        const kids = Array.from(article.children);
        let i = 0;
        while (i < kids.length) {
          const el = kids[i];
          if (!listOf(el)) {
            frag.appendChild(el);
            i++;
            continue;
          }
          const group = [];
          while (i < kids.length && listOf(kids[i])) {
            const cs = getComputedStyle(kids[i]);
            group.push({ el: kids[i], info: listOf(kids[i]), indent: px(cs.marginLeft) + px(cs.paddingLeft) });
            i++;
          }
          const indents = [...new Set(group.map((g) => Math.round(g.indent)))].sort((a, b) => a - b);
          const stack = [];
          let rootList = null;
          for (const g of group) {
            const level = Math.max(g.info.level, indents.indexOf(Math.round(g.indent)));
            const tag = g.info.bullet ? 'UL' : 'OL';
            while (stack.length && stack[stack.length - 1].level > level) stack.pop();
            let top = stack[stack.length - 1];
            if (!top || top.level < level || top.list.tagName !== tag) {
              if (top && top.level === level) stack.pop();
              const list = document.createElement(tag);
              if (tag === 'OL') {
                if (g.info.type !== 'decimal') list.style.listStyleType = g.info.type;
                if (g.info.value !== 1) list.start = g.info.value; // numbering continues from before
              } else {
                const bt = bulletType(g.info.bulletText);
                if (bt !== 'disc' || level > 0) list.style.listStyleType = bt;
              }
              const parent = stack[stack.length - 1];
              if (parent && parent.list.lastElementChild) parent.list.lastElementChild.appendChild(list);
              else if (!rootList || stack.length === 0) {
                frag.appendChild(list);
                rootList = list;
              } else rootList.appendChild(list);
              stack.push({ list, level });
              top = stack[stack.length - 1];
            }
            const item = document.createElement('li');
            while (g.el.firstChild) item.appendChild(g.el.firstChild);
            const s = styles.get(g.el) || {};
            delete s.marginLeft;
            delete s.paddingLeft;
            delete s.textIndent;
            styles.set(item, s);
            top.list.appendChild(item);
          }
        }
        return frag;
      };
      // Notes become small numbered paragraphs under a short line, like Word's footnotes.
      const notesBlock = (ol) => {
        const frag = document.createDocumentFragment();
        const hr = document.createElement('hr');
        hr.style.cssText = 'width:33%;margin:12pt auto 6pt 0;border:0;border-top:1px solid #000;height:0';
        frag.appendChild(hr);
        Array.from(ol.children).forEach((li, i) => {
          let p = li.querySelector(':scope > p');
          if (!p) { p = document.createElement('p'); p.append(...li.childNodes); }
          if (!styles.has(p)) styles.set(p, {});
          const s = styles.get(p);
          ['textIndent', 'marginLeft', 'paddingLeft', 'marginInlineStart'].forEach((k) => delete s[k]);
          // fonts the paragraph took from the list it sat in
          for (const from of [li, ol]) {
            const f = styles.get(from) || {};
            for (const k of INHERITED) if (f[k] && !(k in s)) s[k] = f[k];
          }
          // docx-preview's tab after the note number
          p.querySelectorAll('span').forEach((sp) => { if (!sp.textContent.trim() && sp.style.wordSpacing) sp.remove(); });
          const sup = document.createElement('sup');
          sup.textContent = String((ol.start || 1) + i);
          p.prepend(sup, ' ');
          if (!s.fontSize) s.fontSize = '10pt';
          frag.appendChild(p);
        });
        return frag;
      };
      sections.forEach((sec, idx) => {
        if (idx > 0) {
          const pb = document.createElement('div');
          pb.className = 'page-break';
          out.appendChild(pb);
        }
        const article = sec.querySelector(':scope > article');
        if (article) out.appendChild(convertSection(article));
        // footnotes / endnotes rendered after the article
        Array.from(sec.children).forEach((c) => {
          if (c === article || /^(HEADER|FOOTER)$/.test(c.tagName)) return;
          if (c.tagName === 'OL') out.appendChild(notesBlock(c));
          else out.appendChild(c);
        });
      });

      // 3) Write the snapshots inline, drop class names, tidy markup.
      for (const el of out.querySelectorAll('*')) {
        if (el.namespaceURI === SVGNS || el.closest('math')) continue;
        const s = styles.get(el);
        if (s) applyStyle(el, s);
        if (meta.get(el) && meta.get(el).tab) {
          el.dataset.tab = el.style.textDecorationStyle === 'dotted' ? 'dot' : el.style.textDecorationLine.includes('underline') ? 'line' : 'space';
          el.style.whiteSpace = 'pre';
        }
        if (!el.classList.contains('page-break')) el.removeAttribute('class');
      }
      // outline numbers (1.1, (a), i.) as text at the start of the paragraph
      // (also list items that could not become a real list, e.g. inside table cells)
      for (const [p, r] of numbering) {
        if (!r.label || !out.contains(p)) continue;
        const lab = document.createElement('span');
        lab.dataset.label = '1';
        lab.textContent = r.bullet && /^[\uf000-\uf8ff]/.test(r.label) ? '•' : r.label;
        lab.style.cssText = `display:inline-block;min-width:${Math.round(r.labelWidth)}px;box-sizing:border-box;padding-right:4px;white-space:nowrap;text-indent:0`;
        if (r.labelCss.fontWeight !== r.pCss.fontWeight) lab.style.fontWeight = r.labelCss.fontWeight;
        if (r.labelCss.fontStyle !== r.pCss.fontStyle) lab.style.fontStyle = r.labelCss.fontStyle;
        if (r.labelCss.color !== r.pCss.color) lab.style.color = r.labelCss.color;
        p.insertBefore(lab, p.firstChild);
      }
      // headings
      for (const [el, m] of meta) {
        if (m.heading && out.contains(el)) {
          const h = document.createElement(m.heading);
          h.setAttribute('style', el.getAttribute('style') || '');
          while (el.firstChild) h.appendChild(el.firstChild);
          el.replaceWith(h);
        }
      }
      const topLevel = (n) => {
        while (n && n.parentNode !== out) n = n.parentNode;
        return n;
      };
      // images: unwrap docx-preview's positioning boxes
      for (const [el, m] of meta) {
        if (!out.contains(el) || el.tagName !== 'IMG') continue;
        el.removeAttribute('style');
        if (m.w) el.style.width = m.w + 'px';
        if (m.h) el.style.height = m.h + 'px';
      }
      for (const [el, m] of meta) {
        if (!m.wrapper || !out.contains(el)) continue;
        const img = el.querySelector('img');
        if (!img) continue;
        if (m.wrapNone) {
          // "in front of / behind text": a free-floating picture anchored to its paragraph
          const anchor = topLevel(el);
          const { box, body } = makeBox(true);
          box.style.cssText = `left:${Math.round(m.left)}px;top:0;width:${img.style.width || 'auto'};border:none;padding:0;background:transparent;min-width:0;min-height:0`;
          box.dataset.dy = String(Math.round(m.top));
          const p = document.createElement('p');
          p.style.margin = '0';
          p.appendChild(img);
          body.appendChild(p);
          el.remove();
          if (anchor) anchor.after(box);
          continue;
        }
        if (m.float === 'left' || m.float === 'right') {
          // Word anchors floating pictures to the top of their paragraph
          img.className = m.float === 'left' ? 'wrap-left' : 'wrap-right';
          const p = el.closest('p,li,h1,h2,h3,h4,td');
          el.replaceWith(img);
          if (p && p.firstChild !== img) p.insertBefore(img, p.firstChild);
          continue;
        }
        if (m.block) {
          img.style.display = 'block';
          if (m.align === 'center') img.style.margin = '0 auto';
          if (m.align === 'right') img.style.marginLeft = 'auto';
        }
        el.replaceWith(img);
      }
      // Word text boxes -> DocDoc text boxes (and other drawings as pictures in place)
      let vi = 0;
      for (const d of drawings) {
        const props = vml[vi++] || {};
        if (!out.contains(d.svg)) continue;
        const st = d.style;
        const w = ptToPx(st.width || '') || d.w || 200;
        const h = ptToPx(st.height || '') || d.h || 40;
        const holder = d.svg.parentElement && d.svg.parentElement.tagName === 'DIV' ? d.svg.parentElement : d.svg;
        const anchor = topLevel(holder);
        if (!d.isBox) {
          // a line / arrow / shape: keep the drawing where it was
          const span = document.createElement('span');
          span.className = 'shape';
          span.contentEditable = 'false';
          span.style.cssText = `display:inline-block;width:${Math.round(w)}px;height:${Math.round(h)}px`;
          d.svg.removeAttribute('style');
          d.svg.setAttribute('width', '100%');
          d.svg.setAttribute('height', '100%');
          span.appendChild(d.svg);
          holder.replaceWith(span);
          continue;
        }
        const wrap = props.wrap || 'none';
        const floating = !/^(topAndBottom|square|tight|through)$/i.test(wrap);
        const { box, body } = makeBox(floating);
        for (const fo of d.svg.querySelectorAll('foreignObject')) while (fo.firstChild) body.appendChild(fo.firstChild);
        // position, relative to the text column
        let left = ptToPx(st['margin-left'] || '0');
        if (st['mso-position-horizontal-relative'] === 'page') left -= marginPx;
        const halign = st['mso-position-horizontal'];
        if (halign === 'right') left = contentW - w;
        else if (halign === 'center') left = (contentW - w) / 2;
        const dy = /^(page|margin|top-margin-area)$/.test(st['mso-position-vertical-relative'] || '') ? 0 : ptToPx(st['margin-top'] || '0');
        const border = props.stroked === false ? 'none' : `${props.weight || '0.75pt'} ${/dot/i.test(props.dash) ? 'dotted' : /dash/i.test(props.dash) ? 'dashed' : 'solid'} ${vmlColor(props.stroke)}`;
        box.style.cssText = `width:${Math.round(w)}px;min-height:${Math.round(h)}px;padding:3.6pt 7.2pt;border:${border};background:${props.filled === false ? 'transparent' : vmlColor(props.fill)}${props.round ? ';border-radius:10px' : ''}`;
        if (floating) {
          box.style.left = Math.round(left) + 'px';
          box.style.top = '0px';
          box.dataset.dy = String(Math.round(dy));
          if (+st['z-index'] < 0) box.style.zIndex = '-1'; // behind text
          holder.remove();
          if (anchor) anchor.after(box); else out.appendChild(box);
        } else {
          // text wraps above and below (or around): keep it in the flow after its paragraph
          if (/^(square|tight|through)$/i.test(wrap)) box.classList.add(left > contentW / 2 ? 'wrap-right' : 'wrap-left');
          else if (left > 0) box.style.marginLeft = Math.round(left) + 'px';
          holder.remove();
          if (anchor) {
            if (/^(square|tight|through)$/i.test(wrap)) anchor.before(box); else anchor.after(box);
          } else out.appendChild(box);
        }
      }
      // runs: keep tabs and repeated spaces; drop empty wrappers
      out.querySelectorAll('span').forEach((sp) => {
        if (/\t| {2,}/.test(sp.textContent)) sp.style.whiteSpace = 'pre-wrap';
        if (!sp.attributes.length) sp.replaceWith(...sp.childNodes);
      });
      // empty paragraphs need a <br> to keep their height
      out.querySelectorAll('p,h1,h2,h3,h4,li').forEach((b) => {
        if (!b.textContent.trim() && !b.querySelector('img,br,table,math,svg')) b.appendChild(document.createElement('br'));
      });
      out.querySelectorAll('[style*="min-height"]').forEach((e) => { if (!e.classList.contains('textbox')) e.style.minHeight = ''; });
      return { html: out.innerHTML, page };
    } finally {
      host.remove();
      styleHost.remove();
    }
  }

  window.DocImport = { importDocx };
})();
