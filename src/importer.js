/* High-fidelity .docx import.
 *
 * docx-preview renders the Word file with its real styles (themes, paragraph
 * and character styles, numbering, table styles). We then read the final
 * computed look of every element, write it inline, and rebuild the result as
 * plain editable HTML (real <ul>/<ol> lists, headings, page breaks).
 */
(() => {
  'use strict';

  const PX_TO_MM = 25.4 / 96;
  const INHERITED = ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'color', 'letterSpacing', 'textTransform', 'fontVariantCaps'];
  const CSS_NAME = (p) => p.replace(/[A-Z]/g, (c) => '-' + c.toLowerCase());
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

  // Read what an element looks like, as a plain object of CSS properties.
  function snapshot(el, parentCs, isBlock) {
    const cs = getComputedStyle(el);
    const out = {};
    for (const p of INHERITED) {
      if (!parentCs || cs[p] !== parentCs[p]) out[p] = cs[p];
    }
    const tag = el.tagName;
    if (!TRANSPARENT(cs.backgroundColor)) out.backgroundColor = cs.backgroundColor;
    if (cs.textDecorationLine && cs.textDecorationLine !== 'none') out.textDecorationLine = cs.textDecorationLine;
    if (cs.verticalAlign === 'super' || cs.verticalAlign === 'sub' || (/px$/.test(cs.verticalAlign) && px(cs.verticalAlign))) {
      out.verticalAlign = cs.verticalAlign;
    }
    if (cs.display === 'inline-block' && tag === 'SPAN' && px(cs.width)) {
      out.display = 'inline-block';
      out.width = cs.width;
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
        else if (c.nodeType === 1 && !BLOCKS.has(c.tagName) && c.tagName !== 'TABLE') walk(c);
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

  function listInfo(el) {
    if (headingTag(el.className)) return null; // outline-numbered headings stay headings
    const m = /(?:^|\s)\S+-num-(\d+)-(\d+)(?:\s|$)/.exec(el.className || '');
    if (!m) return null;
    const before = getComputedStyle(el, '::before').content || '';
    let ordered = /counter\(/.test(before);
    let type = 'decimal';
    if (ordered) {
      const t = /counter\([^,)]+,\s*([a-z-]+)\)/.exec(before);
      if (t) type = t[1];
      if (type === 'none' || type === 'disc' || type === 'bullet') ordered = false;
    }
    const cs = getComputedStyle(el);
    return {
      id: m[1],
      level: +m[2],
      indent: px(cs.marginLeft) + px(cs.paddingLeft),
      ordered,
      type,
    };
  }

  function convertSection(article, styles) {
    const frag = document.createDocumentFragment();
    const kids = Array.from(article.children);
    let i = 0;
    while (i < kids.length) {
      const el = kids[i];
      const li = el.tagName === 'P' ? listInfo(el) : null;
      if (!li) {
        frag.appendChild(el);
        i++;
        continue;
      }
      // Gather consecutive list paragraphs and rebuild real nested lists.
      const group = [];
      while (i < kids.length && kids[i].tagName === 'P' && listInfo(kids[i])) {
        group.push({ el: kids[i], info: listInfo(kids[i]) });
        i++;
      }
      const indents = [...new Set(group.map((g) => Math.round(g.info.indent)))].sort((a, b) => a - b);
      const stack = []; // [{list, level}]
      let rootList = null;
      for (const g of group) {
        const level = Math.max(g.info.level, indents.indexOf(Math.round(g.info.indent)));
        const tag = g.info.ordered ? 'OL' : 'UL';
        while (stack.length && stack[stack.length - 1].level > level) stack.pop();
        let top = stack[stack.length - 1];
        if (!top || top.level < level || top.list.tagName !== tag) {
          if (top && top.level === level) stack.pop();
          const list = document.createElement(tag);
          if (g.info.ordered && g.info.type !== 'decimal') list.style.listStyleType = g.info.type;
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
  }

  async function importDocx(bytes) {
    const host = document.createElement('div');
    host.style.cssText = 'position:fixed;left:-20000px;top:0;visibility:hidden;pointer-events:none';
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

      const sections = Array.from(host.querySelectorAll('section.ddimp'));
      if (!sections.length) throw new Error('Nothing to show in this document');

      // Page set-up from the first section.
      const s0 = getComputedStyle(sections[0]);
      const wMm = px(sections[0].style.width ? s0.width : 0) * PX_TO_MM;
      const hMm = px(sections[0].style.minHeight ? s0.minHeight : 0) * PX_TO_MM;
      const page = {
        widthMm: wMm,
        heightMm: hMm,
        marginMm: round(px(s0.paddingLeft) * PX_TO_MM) || 25.4,
      };

      // 1) Snapshot the final look of everything while the original styles still apply.
      const styles = new Map();
      const meta = new Map();
      for (const sec of sections) {
        const article = sec.querySelector(':scope > article') || sec;
        const walk = (el, parentCs) => {
          const isBlock = BLOCKS.has(el.tagName) || el.tagName === 'DIV' && getComputedStyle(el).display === 'block';
          const { cs, out } = snapshot(el, parentCs, isBlock);
          styles.set(el, out);
          if (el.tagName === 'P') meta.set(el, { heading: headingTag(el.className) });
          if (el.tagName === 'IMG') {
            const r = el.getBoundingClientRect();
            meta.set(el, { w: Math.round(r.width), h: Math.round(r.height) });
          }
          if (el.tagName === 'DIV' && el.querySelector('img')) {
            meta.set(el, { float: cs.float, wrapper: true });
          }
          for (const c of el.children) walk(c, cs);
        };
        // Top-level blocks are compared against nothing, so they keep their full look.
        for (const c of article.children) walk(c, null);
      }
      // Paragraph-level fonts: lift shared run formatting to the paragraph.
      for (const [el] of meta) if (el.tagName === 'P') hoist(el, styles);
      // Containers take their fonts from the editor; paragraphs carry their own.
      const DEFAULTS = { letterSpacing: 'normal', textTransform: 'none', fontVariantCaps: 'normal', fontWeight: '400', fontStyle: 'normal' };
      for (const [el, s] of styles) {
        for (const [k, v] of Object.entries(DEFAULTS)) if (s[k] === v) delete s[k];
        if (/^(TABLE|TBODY|THEAD|TR|TD|TH|DIV|OL|UL)$/.test(el.tagName)) HOIST.forEach((p) => delete s[p]);
        if (s.textAlign === 'left') delete s.textAlign;
      }

      // 2) Convert lists while class names are still available.
      const out = document.createElement('div');
      sections.forEach((sec, idx) => {
        if (idx > 0) {
          const pb = document.createElement('div');
          pb.className = 'page-break';
          out.appendChild(pb);
        }
        const article = sec.querySelector(':scope > article');
        if (article) out.appendChild(convertSection(article, styles));
        // footnotes / endnotes rendered after the article
        Array.from(sec.children).forEach((c) => {
          if (c !== article && !/^(HEADER|FOOTER)$/.test(c.tagName)) out.appendChild(c);
        });
      });

      // 3) Write the snapshots inline, drop class names, tidy markup.
      const all = [out, ...out.querySelectorAll('*')];
      for (const el of all) {
        if (el === out) continue;
        const s = styles.get(el);
        if (s) applyStyle(el, s);
        if (!el.classList.contains('page-break')) el.removeAttribute('class');
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
      // images: unwrap docx-preview's positioning boxes
      for (const [el, m] of meta) {
        if (!out.contains(el)) continue;
        if (el.tagName === 'IMG') {
          el.removeAttribute('style');
          if (m.w) el.style.width = m.w + 'px';
          if (m.h) el.style.height = m.h + 'px';
        }
      }
      for (const [el, m] of meta) {
        if (!m.wrapper || !out.contains(el)) continue;
        const img = el.querySelector('img');
        if (!img) continue;
        if (m.float === 'left') img.className = 'wrap-left';
        if (m.float === 'right') img.className = 'wrap-right';
        el.replaceWith(img);
      }
      // runs: keep tabs and repeated spaces; drop empty wrappers
      out.querySelectorAll('span').forEach((sp) => {
        if (/\t| {2,}/.test(sp.textContent)) sp.style.whiteSpace = 'pre-wrap';
        if (!sp.attributes.length) sp.replaceWith(...sp.childNodes);
      });
      // empty paragraphs need a <br> to keep their height
      out.querySelectorAll('p,h1,h2,h3,h4,li').forEach((b) => {
        if (!b.textContent.trim() && !b.querySelector('img,br,table')) b.appendChild(document.createElement('br'));
      });
      out.querySelectorAll('[style*="min-height"]').forEach((e) => (e.style.minHeight = ''));
      return { html: out.innerHTML, page };
    } finally {
      host.remove();
      styleHost.remove();
    }
  }

  window.DocImport = { importDocx };
})();
